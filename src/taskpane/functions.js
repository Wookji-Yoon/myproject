/* global PowerPoint, document, console */
import { getSelectedSlideIndex, getSelectedSlideId } from "./utils.js";

/**
 * 주어진 태그 딕셔너리를 슬라이드에 추가하는 함수
 * @param {Object} userTags 태그 딕셔너리 (key-value 쌍)
 */
async function exportSelectedSlideAsBase64(formData) {
  return new Promise((resolve, reject) => {
    PowerPoint.run(async (context) => {
      // 현재 선택된 슬라이드 인덱스 가져오기
      const selectedSlideIndex = await getSelectedSlideIndex();
      const realSlideIndex = selectedSlideIndex - 1;

      // 선택된 슬라이드 가져오기
      const selectedSlide = context.presentation.slides.getItemAt(realSlideIndex);

      // 슬라이드 내보내기
      const slideExport = selectedSlide.exportAsBase64();

      // 썸네일 저장하기
      const thumbnail = selectedSlide.getImageAsBase64({
        options: {
          height: 100,
        },
      });
      console.log(formData);
      const { title, tags, timestamp } = formData;
      console.log(title, tags, timestamp);
 

      await context.sync();

      // Base64 값 추출
      const slideBase64Value = slideExport.m_value || slideExport;
      const thumbnailBase64Value = thumbnail.m_value;

      resolve({
        slide: slideBase64Value,
        thumbnail: thumbnailBase64Value,
        title: title,
        tags: tags,
        saved_at: timestamp,
      });
    }).catch((error) => {
      reject(error);
    });
  });
}

/**
 * 주어진 base64 객체에서 필요한 데이터를 추출하여 JSON 데이터로 반환
 * @param {Object} base64 base64 객체
 * @returns {Object} JSON 데이터
 */
function createJsonData(base64) {
  return {
    slides: [
      {
        id: new Date().getTime().toString(), // 고유 ID 생성
        base64: base64.slide,
        thumbnail: base64.thumbnail,
        saved_at: base64.saved_at,
        title: base64.title,
        tags: base64.tags,
      },
    ],
  };
}

async function insertAfterSelectedSlide(slides, id) {
  await PowerPoint.run(async function (context) {
    const selectedSlideId = await getSelectedSlideId();
    const slideBase64 = slides.find((slide) => slide.id === id).base64;

    context.presentation.insertSlidesFromBase64(slideBase64, {
      formatting: PowerPoint.InsertSlideFormatting.keepSourceFormatting,
      targetSlideId: selectedSlideId + "#",
    });

    await context.sync();
  });
}

export { exportSelectedSlideAsBase64, insertAfterSelectedSlide, createJsonData };
