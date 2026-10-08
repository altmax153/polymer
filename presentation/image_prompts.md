# gpt-image 프롬프트 (3·12·13·15번 슬라이드)

현재 세션에는 gpt-image 스킬에 필요한 Codex CLI와 ChatGPT 로그인이 없습니다. 이 스킬은 API 키 사용을 금지하므로 이미지를 생성하지 못했고, 아래 프롬프트만 준비했습니다.
생성 방법: 로컬에 스킬을 설치하고 ChatGPT로 로그인한 뒤 아래처럼 실행합니다.
`node gpt-image/scripts/gpt_image.mjs generate --prompt "<프롬프트>" --size 16:9 --out generated-images/<이름>.png`

| # | 슬라이드 | PPT의 현재 상태 | 생성 후 처리 |
|---|---|---|---|
| 1 | 3 연구 배경 | 점선 자리표시자 | 자리표시자를 삭제하고 그 위치(우측)에 삽입 |
| 2 | 12 장치·측정 | 도형으로 그린 격자 모식도 | 선택 사항: 모식도 옆에 넣거나 모식도를 대체 |
| 3 | 13 실험 조건 | 도형으로 그린 측면 모식도 | 선택 사항: 모식도를 대체 |
| 4 | 15 장애물·작업자 | 도형으로 그린 와류 개념도 | 선택 사항: 개념도를 대체 |

12·13·15번은 이해를 위해 그림이 꼭 필요한 슬라이드입니다. 그래서 수치와 명칭이 정확한 도형 모식도를 미리 넣어 두었습니다. AI 이미지는 글자가 틀리게 나올 수 있으므로 텍스트를 넣지 않도록 지시했고, 라벨은 PPT 도형으로 위에 얹는 것을 권합니다.

## 1. 연구 배경 — 대학 화학 연구실
```
A realistic photograph of a university chemistry research laboratory in South Korea: graduate students in lab coats and safety goggles working at a row of chemical fume hoods, many small reagent bottles of different organic solvents on benches, glassware being washed at a sink, bright fluorescent lighting, documentary style, no visible text or logos, 3:4 portrait composition.
```

## 2. 흄후드 정면 — 15개 격자 측정
```
Clean technical illustration of a laboratory chemical fume hood viewed straight from the front, sash fully open, the face opening divided by thin dashed lines into a 5 columns by 3 rows grid, a hot-wire anemometer probe mounted on a tripod stand positioned at the center of one grid cell, flat vector style, white background, navy and teal color palette, no text, no numbers, 16:9.
```

## 3. 실험 조건 — 측면 배치
```
Flat vector cross-section diagram of a laboratory fume hood seen from the side: a vertical sliding glass sash partially lowered, a cardboard box obstacle inside on the work surface, a 1-liter glass beaker of clear liquid placed just inside the sash, a small gas sensor placed outside the sash at chest height, a standing mannequin facing the hood, thin teal arrows showing air flowing into the hood, white background, navy and teal palette, no text, 16:9.
```

## 4. 작업자 앞 와류
```
Top-down conceptual airflow illustration: a fume hood opening drawn as a long horizontal rectangle at the top, a person standing in front of it seen from above, smooth teal streamlines entering the hood at both sides, and a red swirling recirculation vortex forming in the gap between the person's chest and the hood opening, minimal flat vector style, white background, no text, 16:9.
```
