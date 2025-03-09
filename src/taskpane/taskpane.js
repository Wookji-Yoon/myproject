/* global document, Office, console*/

import { signIn, createJsonFile, readJsonFile, fileExists, createFolder } from "./graphService";
import { exportSelectedSlideAsBase64, insertAfterSelectedSlide } from "./functions";
import Tagify from "@yaireo/tagify";

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint && info.platform === Office.PlatformType.PC) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // 네비게이션 초기화
    initializeNavigation();

    // 기본 페이지 표시
    showPage("list-page");

    // Tagify 초기화는 페이지가 표시된 후에 수행
    // 모든 페이지의 이벤트 핸들러 등록
    registerPageEventHandlers("add-page");

    // Tagify 초기화는 더 이상 여기서 하지 않음
    // var input = document.querySelector('input[name=basic]');
    // new Tagify(input);

    // Microsoft Graph 관련 이벤트 핸들러에 tryCatch 적용
    document.getElementById("sign-in").onclick = () =>
      tryCatch(async () => {
        const account = await signIn();
        setMessage(`${account.name}님으로 성공적으로 로그인했습니다!`);
        
        // myapp 폴더 생성 (없는 경우)
        await createFolder();
        
        // JSON 파일 경로 설정
        const jsonFilePath = "/me/drive/root:/myapp/";
        // 파일 존재 여부 확인
        const exists = await fileExists(jsonFilePath + "slides.json");
        // 파일이 없으면 생성
        if (!exists) {
          await createJsonFile(jsonFilePath);
          setMessage(`${account.name}님으로 로그인 완료 및 프레젠테이션 JSON 파일을 생성했습니다.`);
        } else {
          setMessage(`${account.name}님으로 로그인 완료. 프레젠테이션 JSON 파일이 이미 존재합니다.`);
        }
      });

    document.getElementById("add-slide-button").onclick = () =>
      tryCatch(async () => {
        // 폼 데이터 가져오기
        const slideTitle = document.getElementById("slide-title").value;

        // Tagify로 생성된 태그 요소 가져오기
        const tagifyInput = document.querySelector('input[name="basic"]');
        let tags = [];

        try {
          // Tagify 인스턴스가 있는 경우
          if (tagifyInput && tagifyInput.tagify) {
            // tagify의 값을 단순 문자열 배열로 변환
            tags = tagifyInput.tagify.value.map((item) => item.value || "");
          }
          // Tagify 값이 문자열로 존재하는 경우 (JSON 문자열일 수 있음)
          else if (tagifyInput && tagifyInput.value && tagifyInput.value.trim()) {
            try {
              // JSON 형식인지 확인
              if (tagifyInput.value.startsWith("[") && tagifyInput.value.includes("value")) {
                const parsedTags = JSON.parse(tagifyInput.value);
                tags = parsedTags.map((item) => item.value || "");
              } else {
                // 일반 쉼표로 구분된 텍스트
                tags = tagifyInput.value
                  .split(",")
                  .map((tag) => tag.trim())
                  .filter((tag) => tag);
              }
            } catch (e) {
              console.error("태그 파싱 오류:", e);
              // 파싱 오류 시 원본 텍스트를 쉼표로 분리
              tags = tagifyInput.value
                .split(",")
                .map((tag) => tag.trim())
                .filter((tag) => tag);
            }
          }
        } catch (e) {
          console.error("태그 처리 중 오류 발생:", e);
          // 오류 발생 시 빈 배열 유지
          tags = [];
        }

        // 빈 문자열 태그 제거
        tags = tags.filter((tag) => tag);

        // 폼 데이터 객체 생성
        const formData = {
          title: slideTitle,
          tags: tags,
          timestamp: new Date().toISOString(),
        };

        // 콘솔에 데이터 출력
        console.log("슬라이드 추가 폼 데이터:", formData);

        // 슬라이드 추가 폼 데이터를 사용하여 슬라이드 추가
        try {
          const result = await exportSelectedSlideAsBase64(formData);
          console.log("슬라이드 export 성공");
          console.log(result);
        } catch (error) {
          console.error("슬라이드 export 실패:", error);
        }
      });

    // read-json 버튼이 존재하는 경우에만 이벤트 핸들러 등록
    const readJsonButton = document.getElementById("read-json");
    if (readJsonButton) {
      readJsonButton.onclick = () => {
        console.log("read-json 버튼의 onclick이 호출됨");
        tryCatch(async () => {
          // 디버깅 메시지 추가
          console.log("Read JSON 버튼 클릭됨");
          setMessage("JSON 파일 읽기 시작...");
          try {
            // Microsoft Graph 클라이언트 접근 시도로 로그인 상태 확인
            const jsonData = await readJsonFile();
            console.log("읽어온 JSON 데이터:", jsonData);
            // slides 속성이 있는지 확인
            if (jsonData && jsonData.slides && Array.isArray(jsonData.slides)) {
              displaySlides(jsonData.slides);
              setMessage("JSON 파일을 성공적으로 읽었습니다!");
            } else {
              console.error("유효하지 않은 JSON 데이터 형식:", jsonData);
              setMessage("JSON 파일 형식이 올바르지 않습니다.");
            }
          } catch (error) {
            console.error("JSON 읽기 오류:", error);
            if (error.message && error.message.includes("access_token")) {
              setMessage("로그인이 필요합니다. 메인 페이지에서 먼저 로그인해 주세요.");
              showPage("main-page");
            } else {
              setMessage(`오류 발생: ${error.message || error}`);
            }
          }
        });
      };
    }
  }
});

function initializeNavigation() {
  const buttons = document.querySelectorAll(".nav-button");
  buttons.forEach((button) => {
    button.addEventListener("click", () => {
      const pageId = button.dataset.page;
      showPage(pageId);
    });
  });
}

function showPage(pageId) {
  console.log(`페이지 전환: ${pageId}`);
  // 모든 페이지 숨기기
  const pages = document.querySelectorAll(".page");
  pages.forEach((page) => {
    page.style.display = "none";
  });

  // 선택한 페이지만 보이기
  const selectedPage = document.getElementById(pageId);
  if (selectedPage) {
    selectedPage.style.display = "block";
  }

  // 활성 버튼 스타일 변경
  const buttons = document.querySelectorAll(".nav-button");
  buttons.forEach((button) => {
    if (button.dataset.page === pageId) {
      button.classList.add("active");
    } else {
      button.classList.remove("active");
    }
  });

  // 페이지 전환 시 이벤트 핸들러 등록
  registerPageEventHandlers(pageId);
}

// setMessage 함수 추가
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

// clearMessage 함수 추가
async function clearMessage(callback) {
  document.getElementById("message").innerText = "";
  await callback();
}

// tryCatch 함수 추가
async function tryCatch(callback) {
  try {
    document.getElementById("message").innerText = "";
    await callback();
  } catch (error) {
    setMessage("Error: " + error.toString());
  }
}

function displaySlides(slides) {
  const container = document.getElementById("slides-container");
  container.innerHTML = "";

  slides.forEach((slide) => {
    const slideElement = document.createElement("div");
    slideElement.className = "slide-item";

    // Insert 버튼 추가
    const insertButton = document.createElement("button");
    insertButton.className = "insert-button";
    insertButton.textContent = "Insert";
    insertButton.onclick = async () => {
      try {
        await insertAfterSelectedSlide(slides, slide.id);
        setMessage("슬라이드가 성공적으로 삽입되었습니다!");
      } catch (error) {
        setMessage(`슬라이드 삽입 중 오류 발생: ${error.message}`);
      }
    };
    slideElement.appendChild(insertButton);

    // ID 표시
    const idElement = document.createElement("div");
    idElement.className = "slide-info";
    idElement.textContent = `슬라이드 ID: ${slide.id}`;
    slideElement.appendChild(idElement);

    // 썸네일 이미지 표시
    const thumbnailImg = document.createElement("img");
    thumbnailImg.className = "slide-thumbnail";
    thumbnailImg.src = `data:image/png;base64,${slide.thumbnail}`;
    thumbnailImg.alt = "슬라이드 썸네일";
    slideElement.appendChild(thumbnailImg);

    // 태그 표시
    if (slide.tags) {
      const tagsContainer = document.createElement("div");
      tagsContainer.className = "tag-list";

      Object.entries(slide.tags).forEach(([key, value]) => {
        const tagElement = document.createElement("span");
        tagElement.className = "tag-item";
        tagElement.textContent = `${key}: ${value}`;
        tagsContainer.appendChild(tagElement);
      });

      slideElement.appendChild(tagsContainer);
    }

    container.appendChild(slideElement);
  });
}

// 페이지 전환 시 이벤트 핸들러 등록 함수 추가
function registerPageEventHandlers(pageId) {
  console.log(`페이지 ${pageId}의 이벤트 핸들러 등록`);

  if (pageId === "list-page") {
    // 테스트 메시지 버튼
    const testMessageButton = document.getElementById("test-message");
    if (testMessageButton) {
      testMessageButton.onclick = function () {
        console.log("테스트 메시지 버튼 클릭됨");
        setMessage("테스트 메시지가 표시됩니다. 이 메시지가 보이면 메시지 표시 기능이 정상 작동합니다.");
      };

      // Read JSON 버튼 직접 이벤트 등록
      const readJsonButton = document.getElementById("read-json");
      if (readJsonButton) {
        readJsonButton.onclick = function () {
          console.log("read-json 버튼 클릭됨 (페이지 이벤트 핸들러)");
          setMessage("JSON 파일 읽기 시작합니다...");

          // 여기서 실제 JSON 읽기 로직 실행
          tryCatch(async () => {
            try {
              const jsonData = await readJsonFile();
              console.log("JSON 데이터 읽기 성공:", jsonData);

              if (jsonData && jsonData.slides) {
                displaySlides(jsonData.slides);
                setMessage("JSON 파일을 성공적으로 읽었습니다!");
              } else {
                setMessage("JSON 파일 형식이 올바르지 않습니다.");
              }
            } catch (error) {
              console.error("JSON 읽기 오류:", error);
              setMessage(`오류 발생: ${error.message || "알 수 없는 오류"}`);
            }
          });
        };
      }
    }
  } else if (pageId === "add-page") {
    // add-page 페이지에 대한 이벤트 핸들러
    console.log("add-page 이벤트 핸들러 등록");

    // Tagify 초기화
    // 기본 태그 입력 필드
    const basicInput = document.querySelector("input[name=basic]");
    if (basicInput) {
      new Tagify(basicInput, {
        whitelist: ["tag1", "tag2", "tag3", "tag4", "tag5", "tag6", "tag7", "tag8", "tag9", "tag10"],
        dropdown: {
          maxItems: 5,
          classname: "tags-look",
          enabled: 0,
          closeOnSelect: false,
        },
        maxTags: 10,
      });
      console.log("기본 태그 입력 필드에 Tagify 적용됨");
    }

    // 슬라이드 태그 입력 필드
    const slideTagsInput = document.getElementById("slide-tags");
    if (slideTagsInput) {
      new Tagify(slideTagsInput, {
        dropdown: {
          enabled: 0, // 드롭다운 비활성화
        },
      });
      console.log("슬라이드 태그 입력 필드에 Tagify 적용됨");
    }
  }
}

export { setMessage, clearMessage, tryCatch };
