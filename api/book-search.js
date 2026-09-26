// Vercel 서버 함수: /api/book-search?query=책제목
// 카카오(다음) 책 검색 API를 서버에서 대신 호출해서, 브라우저에 API 키가 노출되지 않게 함
// 필요한 환경변수: KAKAO_REST_API_KEY (Vercel 프로젝트 Settings > Environment Variables)
export default async function handler(req, res) {
  const query = String(req.query.query || '').trim();
  if (!query) return res.status(400).json({ error: 'query가 필요해요.' });

  const key = process.env.KAKAO_REST_API_KEY;
  if (!key) return res.status(500).json({ error: 'KAKAO_REST_API_KEY 환경변수가 설정되지 않았어요.' });

  try {
    const r = await fetch(
      `https://dapi.kakao.com/v3/search/book?query=${encodeURIComponent(query)}&size=10`,
      { headers: { Authorization: `KakaoAK ${key}` } }
    );
    if (!r.ok) return res.status(r.status).json({ error: `카카오 검색 실패 (${r.status})` });
    const data = await r.json();

    const results = (data.documents || []).map((d) => {
      // 썸네일 주소 안의 fname 값이 원본(더 큰) 표지 이미지 주소라서, 있으면 그걸 우선 사용
      let cover = d.thumbnail || '';
      try {
        const fname = new URL(cover).searchParams.get('fname');
        if (fname) cover = decodeURIComponent(fname);
      } catch (e) { /* 썸네일 주소가 없거나 형식이 다르면 그대로 사용 */ }
      return {
        title: d.title || '',
        author: (d.authors || []).join(', '),
        publisher: d.publisher || '',
        cover: cover.replace(/^http:/, 'https:'),
      };
    });

    res.setHeader('Cache-Control', 's-maxage=86400'); // 같은 검색어는 하루 동안 캐시
    return res.status(200).json({ results });
  } catch (e) {
    return res.status(500).json({ error: '검색 중 오류가 발생했어요.' });
  }
}
