/* global Office, PowerPoint, console, document, location */

import { getSelectedSlideIndex, getSelectedSlideId, setMessage, tryCatch, formatTagOutput } from "./utils.js";
import {
  signIn,
  createFolder,
  fileExists,
  createJsonFile,
  readJsonFile,
  updateJsonFile,
  editJsonFile,
  deleteOneSlideJsonFile,
  getAccountInfo,
  signOut,
} from "./graphService";
import { getSlideListCache, clearSlideListCache, getSlideCache, addSlideCache, updateSlideListCache } from "./state";

import Tagify from "@yaireo/tagify";
import { v4 as uuidv4 } from "uuid";

async function exportSelectedSlideAsBase64() {
  return new Promise((resolve, reject) => {
    PowerPoint.run(async (context) => {
      const selectedSlideIndex = await getSelectedSlideIndex();
      const realSlideIndex = selectedSlideIndex - 1;

      const selectedSlide = context.presentation.slides.getItemAt(realSlideIndex);

      const thumbnailbase64 = selectedSlide.getImageAsBase64({
        options: {
          height: 100,
        },
      });

      const slidebase64 = selectedSlide.exportAsBase64();

      await context.sync();

      resolve({ thumbnail: thumbnailbase64.m_value, slide: slidebase64.m_value });
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
    const jsonData = await getSlideListCache();

    // 슬라이드 삽입
    await insertAfterSelectedSlide(jsonData.slides, slideId);

    //삽입된 슬라이드로 이동
    goToNextSlide();

    setMessage(`슬라이드 ${slideId}가 성공적으로 삽입되었습니다`);
  } catch (error) {
    console.error("슬라이드 삽입 실패:", error);
    setMessage("슬라이드 삽입에 실패했습니다: " + error.message);
  }
}
function goToNextSlide() {
  Office.context.document.goToByIdAsync(Office.Index.Next, Office.GoToType.Index, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      setMessage("Error: " + asyncResult.error.message);
    }
  });
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

  // 모든 페이지에서 공통으로 사용하는 햄버거 메뉴 이벤트 등록
  const hamburgerMenu = document.getElementById("hamburger-menu");
  if (hamburgerMenu) {
    // 이미 이벤트 리스너가 있는지 확인
    if (!hamburgerMenu.dataset.hasListener) {
      hamburgerMenu.addEventListener("click", function () {
        showPage("settings-page");
      });
      hamburgerMenu.dataset.hasListener = "true";
    }
  }

  // 설정 페이지 뒤로가기 버튼 이벤트 등록
  const settingsBackButton = document.getElementById("settings-page-back");
  if (settingsBackButton) {
    // 이미 이벤트 리스너가 있는지 확인
    if (!settingsBackButton.dataset.hasListener) {
      settingsBackButton.addEventListener("click", function () {
        // 이전 페이지로 돌아가기 (기본적으로 list-page로 설정)
        showPage("list-page");
      });
      settingsBackButton.dataset.hasListener = "true";
    }
  }

  // 계정 페이지 뒤로가기 버튼 이벤트 등록
  const accountBackButton = document.getElementById("account-page-back");
  if (accountBackButton) {
    // 이미 이벤트 리스너가 있는지 확인
    if (!accountBackButton.dataset.hasListener) {
      accountBackButton.addEventListener("click", function () {
        // 설정 페이지로 돌아가기
        showPage("settings-page");
      });
      accountBackButton.dataset.hasListener = "true";
    }
  }

  // 도움말 페이지 뒤로가기 버튼 이벤트 등록
  const helpBackButton = document.getElementById("help-page-back");
  if (helpBackButton) {
    // 이미 이벤트 리스너가 있는지 확인
    if (!helpBackButton.dataset.hasListener) {
      helpBackButton.addEventListener("click", function () {
        // 설정 페이지로 돌아가기
        showPage("settings-page");
      });
      helpBackButton.dataset.hasListener = "true";
    }
  }

  // 업데이트 페이지 뒤로가기 버튼 이벤트 등록
  const updatesBackButton = document.getElementById("updates-page-back");
  if (updatesBackButton) {
    // 이미 이벤트 리스너가 있는지 확인
    if (!updatesBackButton.dataset.hasListener) {
      updatesBackButton.addEventListener("click", function () {
        // 설정 페이지로 돌아가기
        showPage("settings-page");
      });
      updatesBackButton.dataset.hasListener = "true";
    }
  }

  // 설정 항목 클릭 이벤트 등록
  if (pageId === "settings-page") {
    const settingsItems = document.querySelectorAll(".settings-item");
    settingsItems.forEach((item) => {
      // 이미 이벤트 리스너가 있는지 확인
      if (!item.dataset.hasListener) {
        item.addEventListener("click", function () {
          const settingType = this.dataset.setting;
          handleSettingClick(settingType);
        });
        item.dataset.hasListener = "true";
      }
    });
  }

  if (pageId === "list-page") {
    // list-page 페이지에 대한 이벤트 핸들러
    console.log("list-page 이벤트 핸들러 등록");
    tryCatch(async () => {
      try {
        const jsonData = await getSlideListCache();

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

    tryCatch(async () => {
      const searchInput = document.querySelector("input[id=tag-search-input]");

      // 이미 Tagify가 초기화되었는지 확인
      if (!searchInput.tagify) {
        const tagJsonData = await readJsonFile("/me/drive/root:/myapp/tags.json");
        console.log("태그 JSON 데이터 읽기 성공:", tagJsonData);
        searchInput.tagify = new Tagify(searchInput, {
          whitelist: [...new Set([...tagJsonData.tags])],
          dropdown: {
            maxItems: 5,
            classname: "tags-look",
            enabled: 0,
            clearOnSelect: true,
          },
          enforceWhitelist: true,
        });

        // Tagify change 이벤트 리스너 추가
        searchInput.addEventListener("change", function (e) {
          tryCatch(async () => {
            console.log("태그 검색 시도:", e.target.value);
            await handleTagSearch(e.target.value);
          });
        });
      }
    });
  } else if (pageId === "add-page") {
    // add-page 페이지에 대한 이벤트 핸들러
    console.log("add-page 이벤트 핸들러 등록");
    //input 초기화
    const slideTitleInput = document.querySelector("input[id=slide-title]");
    if (slideTitleInput) {
      slideTitleInput.value = "";
    }

    const tagsInput = document.querySelector("input[name=basic]");
    if (tagsInput) {
      tagsInput.value = "";
    }

    tryCatch(async () => {
      const tagJsonData = await readJsonFile("/me/drive/root:/myapp/tags.json");
      console.log("태그 JSON 데이터 읽기 성공:", tagJsonData);

      // Tagify 초기화
      // 기본 태그 입력 필드
      if (tagsInput) {
        const tagify = new Tagify(tagsInput, {
          whitelist: [...new Set([...tagJsonData.tags])],
          dropdown: {
            maxItems: 5,
            classname: "tags-look",
            enabled: 0,
            closeOnSelect: false,
          },
          maxTags: 10,
        });

        // 포커스 이벤트 핸들러 추가
        tagify.DOM.input.addEventListener("focus", function () {
          tagify.DOM.scope.classList.add("tagify--focus");
        });

        tagify.DOM.input.addEventListener("blur", function () {
          tagify.DOM.scope.classList.remove("tagify--focus");
        });

        // 썸네일 이미지 추가
        const { thumbnail, slide } = await exportSelectedSlideAsBase64();

        const thumbnailImg = document.querySelector("#add-slide-thumbnail");
        thumbnailImg.src = `data:image/png;base64,${thumbnail}`;
        thumbnailImg.dataset.slide = slide;
        console.log("썸네일 이미지 설정 완료");

        // 슬라이드 base64 추가
        console.log("기본 태그 입력 필드에 Tagify 적용됨");
      }
    });
  } else if (pageId === "edit-page") {
    // edit-page 페이지에 대한 이벤트 핸들러
    console.log("edit-page 이벤트 핸들러 등록");
    tryCatch(async () => {
      const slideCache = await getSlideCache();
      console.log("슬라이드 데이터 읽기 성공:", slideCache);

      //html의 input의 value에 슬라이드 데이터 추가
      const editSlideTitleInput = document.querySelector("input[id=edit-slide-title]");
      if (editSlideTitleInput) {
        editSlideTitleInput.value = slideCache.title;
      }

      // 썸네일 이미지 추가
      const thumbnailImg = document.querySelector("#edit-slide-thumbnail");
      if (thumbnailImg && slideCache.thumbnail) {
        thumbnailImg.src = `data:image/png;base64,${slideCache.thumbnail}`;
        console.log("썸네일 이미지 설정 완료");
      }

      // 태그 입력 필드에 태그 추가
      const editTagsInput = document.querySelector("input[name=edit-tags]");
      if (editTagsInput) {
        editTagsInput.value = slideCache.tags ? slideCache.tags.join(",") : "";
      }

      const tagJsonData = await readJsonFile("/me/drive/root:/myapp/tags.json");
      console.log("태그 JSON 데이터 읽기 성공:", tagJsonData);

      // Tagify 초기화
      // 수정 페이지 태그 입력 필드
      if (editTagsInput) {
        const editTagify = new Tagify(editTagsInput, {
          whitelist: [...new Set([...tagJsonData.tags])],
          dropdown: {
            maxItems: 5,
            classname: "tags-look",
            enabled: 0,
            closeOnSelect: false,
          },
          maxTags: 10,
        });

        // 포커스 이벤트 핸들러 추가
        editTagify.DOM.input.addEventListener("focus", function () {
          editTagify.DOM.scope.classList.add("tagify--focus");
        });

        editTagify.DOM.input.addEventListener("blur", function () {
          editTagify.DOM.scope.classList.remove("tagify--focus");
        });

        console.log("수정 페이지 태그 입력 필드에 Tagify 적용됨");
      }
    });
  } else if (pageId === "account-page") {
    displayAccountInfo();
  }
}

/**
 * 설정 항목 클릭 처리 함수
 * @param {string} settingType 클릭한 설정 타입
 */
function handleSettingClick(settingType) {
  console.log(`설정 항목 클릭: ${settingType}`);

  switch (settingType) {
    case "account":
      showPage("account-page");
      displayAccountInfo();
      break;
    case "help":
      showPage("help-page");
      break;
    case "updates":
      showPage("updates-page");
      displayUpdateInfo();
      break;
    case "logout":
      handleLogout();
      break;
    default:
      break;
  }
}

/**
 * 계정 정보를 표시하는 함수
 */
async function displayAccountInfo() {
  try {
    // 로딩 상태 표시
    document.getElementById("user-name").textContent = "로딩 중...";
    document.getElementById("user-email").textContent = "로딩 중...";

    // 계정 정보 가져오기
    const accountInfo = await getAccountInfo();
    console.log("계정 정보:", accountInfo);

    // 기본 정보 표시 (이름 및 이메일)
    document.getElementById("user-name").textContent =
      accountInfo.detailedInfo.displayName || accountInfo.basicInfo.name;
    document.getElementById("user-email").textContent = accountInfo.detailedInfo.mail || accountInfo.basicInfo.username;
  } catch (error) {
    console.error("계정 정보 표시 오류:", error);
    document.getElementById("user-name").textContent = "정보를 불러올 수 없습니다";
    document.getElementById("user-email").textContent = "정보를 불러올 수 없습니다";
  }
}

/**
 * 로그아웃 처리 함수
 */
async function handleLogout() {
  try {
    // 로그아웃 시도
    await signOut();

    showPage("main-page");
    clearSlideListCache();
    // 앱 상태 초기화 (필요 시)
  } catch (error) {
    console.error("로그아웃 오류:", error);
    setMessage("로그아웃 중 오류가 발생했습니다: " + error.message);
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

  // 슬라이드 갯수 업데이트
  const slidesCountText = document.getElementById("slides-count-text");
  if (slidesCountText) {
    slidesCountText.textContent = `총 ${slides.length}개 슬라이드`;
  }

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
  // 버튼 로딩 상태 설정
  const addButton = document.getElementById("add-slide-button");
  const originalButtonText = addButton.textContent;
  addButton.textContent = "진행 중...";
  addButton.classList.add("loading");

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

    // 버튼 원래 상태로 복원
    addButton.textContent = originalButtonText;
    addButton.classList.remove("loading");

    return; // 함수 실행 중단
  }

  try {
    // Tagify로 생성된 태그 요소 가져오기
    const tagifyInput = document.querySelector('input[name="basic"]');
    const formattedTags = formatTagOutput(tagifyInput.value);
    const thumbnail = document.querySelector("#add-slide-thumbnail");
    const thumbnailBase64 = thumbnail.src.split(",")[1];
    const slideBase64 = thumbnail.dataset.slide;

    // 폼 데이터 객체 생성
    const formatData = {
      id: uuidv4(),
      title: slideTitle,
      tags: formattedTags,
      saved_at: new Date().toISOString(),
      thumbnail: thumbnailBase64,
      slide: slideBase64,
    };
    // 슬라이드 추가 폼 데이터를 사용하여 슬라이드 추가
    await updateJsonFile(formatData);
    console.log("슬라이드 export 성공");

    // 메시지 표시
    setMessage("슬라이드가 성공적으로 추가되었습니다!");

    // 캐시 초기화 - 새로운 슬라이드가 추가되었으므로
    clearSlideListCache();
    location.reload();
  } catch (error) {
    console.error("슬라이드 내보내기 오류:", error);
    setMessage("슬라이드 내보내기에 실패했습니다: " + error.message);
    // 버튼 원래 상태로 복원
    addButton.textContent = originalButtonText;
    addButton.classList.remove("loading");
  }
}

/**
 * 슬라이드 수정 처리 함수
 * @returns {Promise<void>}
 */
async function handleEditSlide() {
  // 버튼 로딩 상태 설정
  const editButton = document.getElementById("edit-slide-button");
  const originalButtonText = editButton.textContent;
  editButton.textContent = "진행 중...";
  editButton.classList.add("loading");

  // 폼 데이터 가져오기
  const slideTitleInput = document.getElementById("edit-slide-title");
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

    // 버튼 원래 상태로 복원
    editButton.textContent = originalButtonText;
    editButton.classList.remove("loading");

    return; // 함수 실행 중단
  }

  try {
    // Tagify로 생성된 태그 요소 가져오기
    const tagifyInput = document.querySelector('input[name="edit-tags"]');
    const formattedTags = formatTagOutput(tagifyInput.value);

    const updatedSlide = await getSlideCache();
    updatedSlide.title = slideTitle;
    updatedSlide.tags = formattedTags;
    await editJsonFile(updatedSlide);
    console.log("슬라이드 수정 성공");

    // 캐시 초기화 - 새로운 슬라이드가 추가되었으므로
    clearSlideListCache();
    location.reload();

    // 메시지 표시
    setMessage("슬라이드가 성공적으로 수정되었습니다!");

    // 목록 페이지로 이동
    showPage("list-page");
  } catch (error) {
    console.error("슬라이드 수정 오류:", error);
    setMessage("슬라이드 수정에 실패했습니다: " + error.message);
  } finally {
    // 버튼 원래 상태로 복원
    editButton.textContent = originalButtonText;
    editButton.classList.remove("loading");
  }
}

/**
 * 삭제 아이콘 클릭 처리 함수
 * @param {Event} event 클릭 이벤트
 * @returns {Promise<void>}
 */
async function handleDeleteIconClick(event) {
  // 클릭된 요소가 아이콘인 경우 상위 버튼 요소 찾기
  const deleteButton = event.target.classList.contains("delete-icon")
    ? event.target
    : event.target.closest(".delete-icon");
  const slideId = deleteButton.dataset.slideId;

  //프론트엔드에서 슬라이드 삭제
  const slideElement = deleteButton.closest(".slide-item");
  slideElement.remove();

  // 슬라이드 캐시 삭제
  await updateSlideListCache(slideId);

  try {
    // 백엔드에서 슬라이드 삭제 시도
    await deleteOneSlideJsonFile(slideId);
    // 성공 시 최종 메시지 표시
    setMessage(`슬라이드 ID: ${slideId}가 삭제되었습니다`);
  } catch (error) {
    // 실패시 새롭게 SlideList 받아오기 위해 새로고침
    location.reload();
    setMessage(`슬라이드 삭제에 실패했습니다: ${error.message}`);
  }
}

/**
 * 수정 아이콘 클릭 처리 함수
 * @param {Event} event 클릭 이벤트
 * @returns {Promise<void>}
 */
async function handleEditIconClick(event) {
  // 클릭된 요소가 아이콘인 경우 상위 버튼 요소 찾기
  const infoButton = event.target.classList.contains("edit-icon") ? event.target : event.target.closest(".edit-icon");
  const slideId = infoButton.dataset.slideId;

  // 슬라이드 캐시 설정
  await addSlideCache(slideId);
  showPage("edit-page");

  setMessage(`수정 버튼 클릭됨. 슬라이드 ID: ${slideId}`);
}

/**
 * 제목으로 슬라이드를 검색하는 함수
 * @param {string} searchTerm - 검색어
 */
async function handleTitleSearch(searchTerm) {
  try {
    // 캐시된 슬라이드 목록 가져오기
    const cache = await getSlideListCache();
    const slides = cache.slides;

    if (!slides || slides.length === 0) {
      setMessage("검색할 슬라이드가 없습니다.", "warning");
      return;
    }

    // 검색어가 비어있으면 모든 슬라이드 표시
    if (!searchTerm || searchTerm.trim() === "") {
      displaySlides(slides);
      return;
    }

    // 제목에 검색어가 포함된 슬라이드 필터링
    const filteredSlides = slides.filter(
      (slide) => slide.title && slide.title.toLowerCase().includes(searchTerm.toLowerCase())
    );

    if (filteredSlides.length === 0) {
      setMessage(`"${searchTerm}" 검색 결과가 없습니다.`, "info");
      // 검색 결과가 없어도 빈 배열로 표시하여 UI 갱신
      displaySlides([]);
    } else {
      setMessage("", "none"); // 메시지 지우기
      displaySlides(filteredSlides);
    }
  } catch (error) {
    setMessage(`검색 중 오류가 발생했습니다: ${error.message}`, "error");
    console.error("검색 오류:", error);
  }
}

/**
 * 태그로 슬라이드를 검색하는 함수
 * @param {string} searchTerm - 검색할 태그
 */
async function handleTagSearch(searchTerm) {
  const tagSearchButton = document.getElementById("tag-search-button");

  try {
    // 캐시된 슬라이드 목록 가져오기
    const cache = await getSlideListCache();
    const slides = cache.slides;

    const formattedTag = formatTagOutput(searchTerm);

    console.log("태그 검색 시도:", formattedTag);

    // 태그가가 비어있으면 모든 슬라이드 표시
    if (formattedTag.length === 0) {
      tagSearchButton.innerHTML = '<i class="ms-Icon ms-Icon--Filter" aria-hidden="true"></i>';
      displaySlides(slides);
      return;
    }

    // 태그에 검색어가 포함된 슬라이드 필터링
    // 참고로 formattedTag는 string[] 형태이며, slide.tags 또한 string[] 형태이다.
    // formattedTag의 요소가 하나라도 slide.tags에 포함되어 있으면 해당 슬라이드를 표시
    const filteredSlides = slides.filter((slide) => {
      return formattedTag.some((tag) => slide.tags.includes(tag));
    });
    displaySlides(filteredSlides);
    tagSearchButton.innerHTML = '<i class="ms-Icon ms-Icon--FilterSolid" aria-hidden="true"></i>';
  } catch (error) {
    setMessage(`태그 검색 중 오류가 발생했습니다: ${error.message}`, "error");
    console.error("태그 검색 오류:", error);
  }
}

/**
 * 업데이트 정보 가져오기
 * @returns {Object} 업데이트 정보 객체
 */
function getUpdateInfo() {
  // 하드코딩된 업데이트 정보 반환
  // 실제 애플리케이션에서는 이 부분을 서버 API 호출로 대체해야 합니다
  const updateData = {
    current_version: "1.0",
    is_latest: false,
    download_url: "https://www.google.com",
    update_history: [
      {
        version: "1.0",
        release_date: "2023-07-01",
        details: [
          "초기 버전 출시",
          "PowerPoint 슬라이드 저장 및 검색 기능",
          "태그 기반 슬라이드 관리",
          "Microsoft 계정 연동",
        ],
      },
      {
        version: "0.9 Beta",
        release_date: "2023-06-15",
        details: ["베타 버전 릴리스", "기본 UI 구현", "OneDrive 연동 테스트", "슬라이드 썸네일 생성 기능"],
      },
      {
        version: "0.8 Alpha",
        release_date: "2023-05-20",
        details: ["알파 버전 테스트", "기본 기능 구현", "Office API 연동 시작"],
      },
    ],
  };

  return updateData;
}

/**
 * 업데이트 정보 표시 함수
 */
function displayUpdateInfo() {
  try {
    // 로딩 상태 표시
    document.querySelector(".version-number").textContent = "로딩 중...";
    document.querySelector(".version-status").textContent = "";

    // 다운로드 링크 초기화
    const downloadLink = document.getElementById("download-link");
    if (downloadLink) {
      downloadLink.style.display = "none";
    }

    // 업데이트 히스토리 초기화
    const updateHistory = document.querySelector(".update-history");
    updateHistory.innerHTML = "";

    // 업데이트 정보 가져오기
    const updateInfo = getUpdateInfo();
    console.log("업데이트 정보:", updateInfo);

    // 현재 버전 표시
    document.querySelector(".version-number").textContent = `ver. ${updateInfo.current_version}`;

    // 최신 버전 상태 표시
    const versionStatus = document.querySelector(".version-status");
    if (updateInfo.is_latest) {
      versionStatus.textContent = "최신버전입니다";
      versionStatus.classList.remove("outdated");

      // 다운로드 링크 숨기기
      if (downloadLink) {
        downloadLink.style.display = "none";
      }
    } else {
      versionStatus.textContent = "최신버전이 아닙니다";
      versionStatus.classList.add("outdated");

      // 다운로드 링크 표시
      if (downloadLink && updateInfo.download_url) {
        downloadLink.href = updateInfo.download_url;
        downloadLink.style.display = "block";
      }
    }

    // 업데이트 히스토리 표시
    updateInfo.update_history.forEach(function (update) {
      const updateItem = document.createElement("div");
      updateItem.className = "update-item";

      const versionElement = document.createElement("div");
      versionElement.className = "update-version";
      versionElement.textContent = `ver. ${update.version}`;

      const detailsList = document.createElement("ul");
      detailsList.className = "update-details";

      update.details.forEach(function (detail) {
        const detailItem = document.createElement("li");
        detailItem.textContent = detail;
        detailsList.appendChild(detailItem);
      });

      updateItem.appendChild(versionElement);
      updateItem.appendChild(detailsList);
      updateHistory.appendChild(updateItem);
    });
  } catch (error) {
    console.error("업데이트 정보 표시 오류:", error);
    document.querySelector(".version-number").textContent = "ver. 1.0";
    document.querySelector(".version-status").textContent = "업데이트 정보를 불러올 수 없습니다";
    document.querySelector(".version-status").style.color = "#ff9800";
  }
}

export {
  exportSelectedSlideAsBase64,
  insertAfterSelectedSlide,
  handleInsertSlide,
  showPage,
  registerPageEventHandlers,
  handleSignIn,
  displaySlides,
  handleExportSlide,
  handleEditSlide,
  handleDeleteIconClick,
  handleEditIconClick,
  handleTitleSearch,
  handleTagSearch,
  handleSettingClick,
  displayAccountInfo,
  handleLogout,
  displayUpdateInfo,
};
