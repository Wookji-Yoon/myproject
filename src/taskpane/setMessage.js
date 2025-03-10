/**
 * 화면에 메시지를 표시하는 함수
 * @param {string} message 표시할 메시지
 */
export function setMessage(message) {
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
