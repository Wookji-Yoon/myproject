/* global PowerPoint, console, document, location */
import { getSelectedSlideIndex, getSelectedSlideId, setMessage, tryCatch } from "./utils.js";
import { signIn, createFolder, fileExists, createJsonFile, readJsonFile, updateJsonFile } from "./graphService";
import Tagify from "@yaireo/tagify";

// 슬라이드 데이터 캐시
let slidesCache = null;

/**
 * 슬라이드 JSON 데이터를 가져오는 함수 (캐시 활용)
 * @returns {Promise<Object>} 슬라이드 데이터
 */
async function getSlidesData() {
  if (slidesCache === null) {
    console.log("슬라이드 데이터를 캐시에서 찾을 수 없음, API 호출");
    const jsonData = await readJsonFile();
    if (jsonData && jsonData.slides) {
      slidesCache = jsonData;
    }
  } else {
    console.log("캐시된 슬라이드 데이터 사용");
  }
  return slidesCache;
}

/**
 * 슬라이드 캐시를 초기화하는 함수
 */
function clearSlidesCache() {
  slidesCache = null;
  console.log("슬라이드 캐시가 초기화되었습니다.");
}

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
        id: new Date().getTime().toString(), // 고유 ID 생성
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

async function insertAfterSelectedSlide(slides, id) {
  await PowerPoint.run(async function (context) {
    const selectedSlideId = await getSelectedSlideId();
    const sourceSlideBase64 = slides.find((slide) => slide.id === id).slide;

    context.presentation.insertSlidesFromBase64(sourceSlideBase64, {
      formatting: PowerPoint.InsertSlideFormatting.keepSourceFormatting,
      targetSlideId: selectedSlideId + "#",
    });

    await context.sync();
  });
}

/**
 * 선택한 슬라이드 ID를 삽입하는 핸들러 함수
 * @param {string} slideId 삽입할 슬라이드 ID
 */
async function handleInsertSlide(slideId) {
  try {
    // 메시지 표시
    setMessage(`슬라이드 ${slideId}를 선택했습니다`);

    // 캐시된 데이터 가져오기
    const jsonData = await getSlidesData();

    // 슬라이드 삽입
    await insertAfterSelectedSlide(jsonData.slides, slideId);

    setMessage(`슬라이드 ${slideId}가 성공적으로 삽입되었습니다`);
  } catch (error) {
    console.error("슬라이드 삽입 실패:", error);
    setMessage("슬라이드 삽입에 실패했습니다: " + error.message);
  }
}

/**
 * 특정 페이지를 표시하는 함수
 * @param {string} pageId 표시할 페이지의 ID
 */
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

/**
 * 페이지 전환 시 이벤트 핸들러 등록 함수
 * @param {string} pageId 이벤트 핸들러를 등록할 페이지의 ID
 */
function registerPageEventHandlers(pageId) {
  console.log(`페이지 ${pageId}의 이벤트 핸들러 등록`);

  if (pageId === "list-page") {
    // list-page 페이지에 대한 이벤트 핸들러
    console.log("list-page 이벤트 핸들러 등록");
    tryCatch(async () => {
      try {
        const jsonData = await getSlidesData();

        if (jsonData && jsonData.slides) {
          displaySlides(jsonData.slides);
          setMessage("JSON 파일을 성공적으로 읽었습니다!");
        } else {
          setMessage("JSON 파일 형식이 올바르지 않습니다.");
        }
      } catch (error) {
        console.error("JSON 읽기 오류:", error);
        setMessage("JSON 파일 읽기 중 오류가 발생했습니다: " + error.message);
      }
    });
  } else if (pageId === "add-page") {
    // add-page 페이지에 대한 이벤트 핸들러
    console.log("add-page 이벤트 핸들러 등록");
    tryCatch(async () => {
      const tagJsonData = await readJsonFile("/me/drive/root:/myapp/tags.json");
      console.log("태그 JSON 데이터 읽기 성공:", tagJsonData);

      // Tagify 초기화
      // 기본 태그 입력 필드
      const basicInput = document.querySelector("input[name=basic]");
      if (basicInput) {
        new Tagify(basicInput, {
          whitelist: [...new Set([...tagJsonData.tags])],
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
    });
  }
}

/**
 * 로그인 처리 및 초기 설정을 수행하는 함수
 * @returns {Promise<void>}
 */
async function handleSignIn() {
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
    showPage("list-page");
  } else {
    setMessage(`${account.name}님으로 로그인 완료. 프레젠테이션 JSON 파일이 이미 존재합니다.`);
    showPage("list-page");
  }
}

/**
 * 슬라이드 목록을 화면에 표시하는 함수
 * @param {Array} slides 표시할 슬라이드 배열
 */
function displaySlides(slides) {
  const container = document.getElementById("slides-container");
  container.innerHTML = "";

  slides.forEach((slide) => {
    const slideElement = document.createElement("div");
    slideElement.className = "slide-item";

    // 썸네일 이미지 표시
    const thumbnailContainer = document.createElement("div");
    thumbnailContainer.className = "thumbnail-container";
    const thumbnailImg = document.createElement("img");
    thumbnailImg.className = "slide-thumbnail";
    thumbnailImg.src = `data:image/png;base64,${slide.thumbnail}`;
    thumbnailImg.alt = "슬라이드 썸네일" + slide.id;
    thumbnailImg.title = "슬라이드 삽입"; // 툴팁 추가
    thumbnailContainer.appendChild(thumbnailImg);

    // 아이콘 컨테이너 추가
    const iconsContainer = document.createElement("div");
    iconsContainer.className = "slide-icons-container";

    // 정보 버튼 추가
    const infoButton = document.createElement("div");
    infoButton.className = "slide-icon edit-icon";
    infoButton.innerHTML = '<i class="ms-Icon ms-Icon--Edit" aria-hidden="true"></i>';
    infoButton.title = "슬라이드 수정";
    infoButton.dataset.slideId = slide.id;

    // 삭제 버튼 추가
    const deleteButton = document.createElement("div");
    deleteButton.className = "slide-icon delete-icon";
    deleteButton.innerHTML = '<i class="ms-Icon ms-Icon--Delete" aria-hidden="true"></i>';
    deleteButton.title = "슬라이드 삭제";
    deleteButton.dataset.slideId = slide.id;

    // 아이콘들을 컨테이너에 추가
    iconsContainer.appendChild(infoButton);
    iconsContainer.appendChild(deleteButton);

    // 아이콘 컨테이너를 썸네일 컨테이너에 추가
    thumbnailContainer.appendChild(iconsContainer);

    // 토글 아이콘 추가
    const toggleIcon = document.createElement("div");
    toggleIcon.className = "toggle-icon";
    toggleIcon.textContent = "▼";
    toggleIcon.title = "제목 숨기기"; // 초기 툴팁 설정
    thumbnailContainer.appendChild(toggleIcon);

    slideElement.appendChild(thumbnailContainer);

    // 슬라이드 정보 표시
    const slideInfo = document.createElement("div");
    slideInfo.className = "slide-info";
    const slideTitle = document.createElement("p");

    // 제목과 날짜 추가
    slideTitle.innerHTML = `<strong>${slide.title}</strong><span>${new Date(slide.saved_at).toLocaleDateString()}</span>`;
    slideInfo.appendChild(slideTitle);
    slideElement.appendChild(slideInfo);

    // 태그 목록 추가
    const tagList = document.createElement("div");
    tagList.className = "tag-list";

    if (slide.tags && Array.isArray(slide.tags)) {
      slide.tags.forEach((tag) => {
        const tagSpan = document.createElement("span");
        tagSpan.className = "tag-item";
        tagSpan.textContent = tag;
        tagList.appendChild(tagSpan);
      });
    }

    slideElement.appendChild(tagList);

    // 컨테이너에 슬라이드 요소 추가
    container.appendChild(slideElement);
  });
}

/**
 * 슬라이드 추가 및 내보내기 처리 함수
 * @returns {Promise<void>}
 */
async function handleExportSlide() {
  // 폼 데이터 가져오기
  const slideTitleInput = document.getElementById("slide-title");
  const slideTitle = slideTitleInput.value.trim();

  // 오류 메시지 요소 제거 (기존에 있을 경우)
  const existingError = document.getElementById("title-error-message");
  if (existingError) {
    existingError.remove();
  }

  // 제목이 비어있는 경우
  if (!slideTitle) {
    // 오류 메시지 생성 및 표시
    const errorMessage = document.createElement("div");
    errorMessage.id = "title-error-message";
    errorMessage.style.color = "red";
    errorMessage.style.fontSize = "12px";
    errorMessage.style.marginTop = "4px";
    errorMessage.textContent = "제목을 입력해주세요";

    // 오류 메시지를 제목 입력 필드 뒤에 삽입
    slideTitleInput.parentNode.insertBefore(errorMessage, slideTitleInput.nextSibling);

    // 제목 입력 필드에 포커스 설정
    slideTitleInput.focus();
    return; // 함수 실행 중단
  }

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
    await updateJsonFile(result);
    console.log("슬라이드 export 성공");
    console.log(result);

    // 캐시 초기화 - 새로운 슬라이드가 추가되었으므로
    clearSlidesCache();

    location.reload();
  } catch (error) {
    console.error("슬라이드 export 실패:", error);
    setMessage("슬라이드 내보내기에 실패했습니다: " + error.message);
  }
}

export {
  exportSelectedSlideAsBase64,
  insertAfterSelectedSlide,
  handleSignIn,
  handleExportSlide,
  setMessage,
  showPage,
  registerPageEventHandlers,
  tryCatch,
  displaySlides,
  handleInsertSlide,
  getSlidesData,
  clearSlidesCache,
};
