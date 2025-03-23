/* global Office, OfficeExtension, console, document, fetch */

function getSelectedSlideIndex() {
  return new OfficeExtension.Promise(function (resolve, reject) {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
      try {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          reject(console.error(asyncResult.error.message));
        } else {
          resolve(asyncResult.value.slides[0].index);
        }
      } catch (error) {
        reject(console.log(error));
      }
    });
  });
}

function getSelectedSlideId() {
  return new OfficeExtension.Promise(function (resolve, reject) {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
      try {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          reject(console.error(asyncResult.error.message));
        } else {
          resolve(asyncResult.value.slides[0].id);
        }
      } catch (error) {
        reject(console.log(error));
      }
    });
  });
}

/**
 * 화면에 메시지를 표시하는 함수
 * @param {string} message 표시할 메시지
 */
function setMessage(message) {
  console.log("메시지 설정:", message);
  const messageElement = document.getElementById("message");
  if (messageElement) {
    messageElement.innerText = message;
  } else {
    console.error("message 요소를 찾을 수 없음");
  }
}

/**
 * 화면에 표시된 메시지를 지우는 함수
 * @param {Function} callback 메시지를 지운 후 실행할 콜백 함수
 * @returns {Promise<void>}
 */

async function clearMessage(callback) {
  document.getElementById("message").innerText = "";
  if (callback) {
    callback();
  }
}

/**
 * 에러 핸들링을 위한 유틸리티 함수
 * @param {Function} callback 실행할 콜백 함수
 */
async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    console.error("Error:", error);
  }
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

/**
 * 배열 A에서 배열 B에 포함된 요소를 개수만큼 제거하는 함수
 *
 * @param {Array} A - 원본 배열
 * @param {Array} B - 제거할 요소가 들어있는 배열
 * @returns {Array} B의 요소를 제거한 A 배열
 */
function subtractArrays(A, B) {
  B.forEach((item) => {
    let index = A.indexOf(item);
    if (index !== -1) A.splice(index, 1); // A에서 하나만 삭제
  });
  return A;
}

function formatTagOutput(output) {
  if (output === "") {
    return "";
  }
  const parsed = JSON.parse(output);
  return parsed.map((item) => item.value);
}

/**
 * 태그 목록에서 특정 키워드를 감지하여 해당하는 이미지를 모달로 표시하는 함수
 * @param {Array} tags - 감지할 태그 배열
 * @param {Object} keywordImageMap - 키워드와 이미지 파일명의 매핑 객체 (예: {'채운': 'logo-filled.png'})
 * @param {Set} detectedKeywords - 이미 감지된 키워드를 추적하는 Set 객체
 * @returns {Set} - 현재까지 감지된 키워드가 업데이트된 Set 객체
 */
function detectKeywordsAndShowImages(tags, keywordImageMap, detectedKeywords = new Set()) {
  if (!Array.isArray(tags) || tags.length === 0) {
    return detectedKeywords;
  }

  // 키워드 검사
  Object.keys(keywordImageMap).forEach((keyword) => {
    // 해당 키워드가 태그에 있고, 아직 감지되지 않았다면
    if (tags.includes(keyword) && !detectedKeywords.has(keyword)) {
      // 이미지 모달 표시
      showImageModal(keywordImageMap[keyword]);
      console.log(`${keyword}가 있습니다`);
      // 감지된 키워드 추가
      detectedKeywords.add(keyword);
    }
  });

  return detectedKeywords;
}

/**
 * 이미지를 모달 형태로 화면 중앙에 표시하는 함수
 * @param {string} imageFileName - 표시할 이미지 파일 이름
 */
function showImageModal(imageFileName) {
  // 이미 모달이 존재하는 경우 제거
  const existingModal = document.getElementById("image-modal-container");
  if (existingModal) {
    document.body.removeChild(existingModal);
    return;
  }

  // 모달 컨테이너 생성
  const modalContainer = document.createElement("div");
  modalContainer.id = "image-modal-container";
  modalContainer.style.position = "fixed";
  modalContainer.style.top = "0";
  modalContainer.style.left = "0";
  modalContainer.style.width = "100%";
  modalContainer.style.height = "100%";
  modalContainer.style.backgroundColor = "rgba(0, 0, 0, 0.7)";
  modalContainer.style.display = "flex";
  modalContainer.style.justifyContent = "center";
  modalContainer.style.alignItems = "center";
  modalContainer.style.zIndex = "9999";

  // 이미지 요소 생성
  const imageElement = document.createElement("img");
  imageElement.src = `https://wookji-yoon.github.io/myproject/assets/${imageFileName}`;
  imageElement.style.maxWidth = "80%";
  imageElement.style.maxHeight = "80%";
  imageElement.style.boxShadow = "0 0 20px rgba(0, 0, 0, 0.5)";
  imageElement.style.transition = "transform 0.2s ease-in-out";

  // 이미지에 마우스 오버 효과 추가
  imageElement.onmouseover = function () {
    this.style.transform = "scale(1.02)";
  };
  imageElement.onmouseout = function () {
    this.style.transform = "scale(1)";
  };

  // 이미지 클릭 시 이벤트 버블링 방지
  imageElement.onclick = function (event) {
    event.stopPropagation();
  };

  // 모달 컨테이너에 이미지 추가
  modalContainer.appendChild(imageElement);

  // 모달 클릭 시 닫기
  modalContainer.addEventListener("click", function () {
    document.body.removeChild(modalContainer);
  });

  // 모달을 body에 추가
  document.body.appendChild(modalContainer);

  // ESC 키를 눌러 모달 닫기 기능 추가
  function handleEscKey(event) {
    if (event.key === "Escape") {
      if (document.body.contains(modalContainer)) {
        document.body.removeChild(modalContainer);
      }
      document.removeEventListener("keydown", handleEscKey);
    }
  }
  document.addEventListener("keydown", handleEscKey);
}

async function checkForUpdates() {
  // API에서 JSON 데이터 가져오기
  const response = await fetch("https://raw.githubusercontent.com/Wookji-Yoon/SliderAPI/refs/heads/master/update.json");
  const data = await response.json();

  console.log(data);

  //data의 형태는 {version: string, update_url: string, update_description: string[]}이다.
  const current_version = "1.1";
  const update_url = data.update_url;
  const update_description = data.update_description;
  if (current_version !== data.version) {
    return { update: true, update_url: update_url, update_description: update_description };
  } else {
    return { update: false };
  }
}

export {
  getSelectedSlideIndex,
  getSelectedSlideId,
  setMessage,
  clearMessage,
  tryCatch,
  createJsonData,
  subtractArrays,
  formatTagOutput,
  detectKeywordsAndShowImages,
  showImageModal,
  checkForUpdates,
};
