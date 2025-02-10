/* global PowerPoint, */
import { getSelectedSlideIndex, getSelectedSlideId } from "./utils.js";

/**
 * 특정 슬라이드에 태그를 추가하는 비동기 함수
 * 태그 키를 대문자로 변환하여 추가
 * @param {string} key 태그 키 (대문자로 변환됨)
 * @param {string} value 태그 값
 * @async
 */
async function addSlideTag(key, value) {
  await PowerPoint.run(async function (context) {
    let selectedSlideIndex = await getSelectedSlideIndex();
    const realSlideIndex = selectedSlideIndex - 1;

    const slide = context.presentation.slides.getItemAt(realSlideIndex);

    // 키를 대문자로 변환하여 태그 추가
    slide.tags.add(key.toUpperCase(), value);

    await context.sync();
  });
}

let chosenFileBase64;

async function exportSelectedSlideAsBase64() {
  return new Promise((resolve, reject) => {
    PowerPoint.run(async (context) => {
      // 현재 선택된 슬라이드 인덱스 가져오기
      const selectedSlideIndex = await getSelectedSlideIndex();
      const realSlideIndex = selectedSlideIndex - 1;

      // 선택된 슬라이드 가져오기
      const selectedSlide = context.presentation.slides.getItemAt(realSlideIndex);

      // 슬라이드 내보내기
      const slideExport = selectedSlide.exportAsBase64();

      await context.sync();

      // Base64 값 추출
      const base64Value = slideExport.m_value || slideExport;

      chosenFileBase64 = base64Value;

      resolve(base64Value);
    }).catch((error) => {
      reject(error);
    });
  });
}

async function insertAfterSelectedSlide() {
  await PowerPoint.run(async function (context) {
    const selectedSlideId = await getSelectedSlideId();

    context.presentation.insertSlidesFromBase64(chosenFileBase64, {
      formatting: PowerPoint.InsertSlideFormatting.keepSourceFormatting,
      targetSlideId: selectedSlideId + "#",
    });

    await context.sync();
  });
}

export { addSlideTag, exportSelectedSlideAsBase64, insertAfterSelectedSlide };
