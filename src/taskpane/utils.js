/* global Office, OfficeExtension, console, document */

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
    // 메시지가 설정되면 스크롤을 최상단으로 이동하여 메시지가 보이도록 함
    messageElement.scrollIntoView({ behavior: "smooth", block: "start" });
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
} /**
 * 에러 핸들링을 위한 유틸리티 함수
 * @param {Function} callback 실행할 콜백 함수
 */
async function tryCatch(callback) {
  try {
    document.getElementById("message").innerText = "";
    await callback();
  } catch (error) {
    setMessage("Error: " + error.toString());
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

export {
  getSelectedSlideIndex,
  getSelectedSlideId,
  setMessage,
  clearMessage,
  tryCatch,
  createJsonData,
  subtractArrays,
};
