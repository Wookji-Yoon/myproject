/* global console */

import { readJsonFile } from "./graphService";

// 슬라이드리스트 데이터 캐시
let slideListCache = null;
let slideCache = null;

/**
 * 슬라이드 JSON 데이터를 가져오는 함수 (캐시 활용)
 * @returns {Promise<Object>} 슬라이드 데이터
 */
async function getSlideListCache() {
  if (slideListCache === null) {
    console.log("슬라이드 데이터를 캐시에서 찾을 수 없음, API 호출");
    const jsonData = await readJsonFile();
    if (jsonData && jsonData.slides) {
      slideListCache = jsonData;
    }
  } else {
    console.log("캐시된 슬라이드 데이터 사용");
  }
  return slideListCache;
}

/**
 * 슬라이드목록 캐시를 업데이트하는 함수
 */
async function updateSlideListCache(slideId) {
  const slideList = await getSlideListCache();
  const updatedSlideList = slideList.slides.filter((slide) => slide.id !== slideId);
  slideListCache = updatedSlideList;
  return slideListCache;
}

/**
 * 슬라이드목록 캐시를 초기화하는 함수
 */
function clearSlideListCache() {
  slideListCache = null;
  console.log("슬라이드 캐시가 초기화되었습니다.");
}

/**
 * 슬라이드 캐시를 설정하는 함수
 */
async function addSlideCache(slideId) {
  const slideList = await getSlideListCache();
  slideCache = slideList.slides.find((slide) => slide.id === slideId);
  return slideCache;
}

/**
 * 슬라이드 캐시를 가져오는 함수
 */
async function getSlideCache() {
  return slideCache;
}

/**
 * 슬라이드 캐시를 초기화하는 함수
 */
async function clearSlideCache() {
  slideCache = null;
  console.log("슬라이드 캐시가 초기화되었습니다.");
}

export { getSlideListCache, clearSlideListCache, addSlideCache, getSlideCache, clearSlideCache, updateSlideListCache };
