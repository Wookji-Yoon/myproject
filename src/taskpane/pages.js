/* global console, document, window */

import { tryCatch, formatTagOutput, detectKeywordsAndShowImages, setMessage, checkForUpdates } from "./utils.js";
import { getSlideListCache, getSlideCache, getTags, clearSlideListCache } from "./state.js";
import Tagify from "@yaireo/tagify";
import { exportSelectedSlideAsBase64 } from "./functions.js";
import { displaySlides, handleTagSearch } from "./functions.js";
import { getAccountInfo, signOut } from "./graphService.js";

/**
 * 헤더 가시성을 업데이트하는 함수
 * @param {string} pageId 표시할 페이지의 ID
 */
function updateHeaderVisibility(pageId) {
  const header = document.querySelector(".app-header");
  const hiddenHeaderPages = ["main-page", "updates-page", "help-page-main"];

  if (header) {
    header.style.display = hiddenHeaderPages.includes(pageId) ? "none" : "flex";
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

  // 헤더 가시성 업데이트
  updateHeaderVisibility(pageId);

  // 페이지 전환 시 이벤트 핸들러 등록
  registerPageEventHandlers(pageId);
}

/**
 * 페이지 전환 시 이벤트 핸들러 등록 함수
 * @param {string} pageId 이벤트 핸들러를 등록할 페이지의 ID
 */
function registerPageEventHandlers(pageId) {
  console.log(`페이지 ${pageId}의 이벤트 핸들러 등록`);

  // 공통 이벤트 핸들러 등록
  registerCommonHandlers();

  // 페이지별 이벤트 핸들러 등록
  switch (pageId) {
    case "main-page":
      registerMainPageHandlers();
      break;
    case "list-page":
      registerListPageHandlers();
      break;
    case "add-page":
      registerAddPageHandlers();
      break;
    case "edit-page":
      registerEditPageHandlers();
      break;
    case "account-page":
      registerAccountPageHandlers();
      break;
    case "settings-page":
      registerSettingsPageHandlers();
      break;
    case "help-page-main":
      registerHelpPageMainHandlers();
      break;
    case "help-page-setting":
      registerHelpPageSettingHandlers();
      break;
    case "version-page":
      registerVersionPageHandlers();
      break;
    case "updates-page":
      registerUpdatesPageHandlers();
      break;
    default:
      console.log(`알 수 없는 페이지 ID: ${pageId}`);
  }
}

/**
 * 메인 페이지 이벤트 핸들러 등록
 */
function registerMainPageHandlers() {
  console.log("메인 페이지 이벤트 핸들러 등록");
}

/**
 * 리스트 페이지 이벤트 핸들러 등록
 */
function registerListPageHandlers() {
  console.log("리스트 페이지 이벤트 핸들러 등록");

  let jsonData;
  tryCatch(async () => {
    const slidesCountText = document.getElementById("slides-count-text");
    slidesCountText.textContent = "로딩 중...";
    try {
      jsonData = await getSlideListCache();

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

    if (!searchInput.tagify) {
      const tags = await getTags();

      searchInput.tagify = new Tagify(searchInput, {
        whitelist: tags,
        dropdown: {
          maxItems: 5,
          classname: "tags-look",
          enabled: 0,
          clearOnSelect: true,
        },
        enforceWhitelist: true,
      });

      searchInput.addEventListener("change", function (e) {
        tryCatch(async () => {
          console.log("태그 검색 시도:", e.target.value);
          await handleTagSearch(e.target.value);
        });
      });
    }
  });
}

/**
 * 추가 페이지 이벤트 핸들러 등록
 */
function registerAddPageHandlers() {
  console.log("추가 페이지 이벤트 핸들러 등록");

  // input 초기화
  const slideTitleInput = document.querySelector("input[id=slide-title]");
  if (slideTitleInput) {
    slideTitleInput.value = "";
  }

  const tagsInput = document.querySelector("input[name=basic]");
  if (tagsInput) {
    tagsInput.value = "";
  }

  tryCatch(async () => {
    const tags = await getTags();

    if (tagsInput) {
      const tagify = new Tagify(tagsInput, {
        whitelist: tags,
        dropdown: {
          maxItems: 5,
          classname: "tags-look",
          enabled: 0,
          closeOnSelect: false,
        },
        maxTags: 10,
      });

      let detectedKeywords = new Set();

      const keywordImageMap = {
        채운: "flower.jpg",
        집가고싶다: "house.png",
        꼬질꼬질: "ggojil.png",
        사랑해: "love.png",
        움하하: "haha.png",
      };

      tagsInput.addEventListener("change", function (e) {
        const formattedTags = formatTagOutput(e.target.value);
        detectedKeywords = detectKeywordsAndShowImages(formattedTags, keywordImageMap, detectedKeywords);
      });

      tagify.DOM.input.addEventListener("focus", function () {
        tagify.DOM.scope.classList.add("tagify--focus");
      });

      tagify.DOM.input.addEventListener("blur", function () {
        tagify.DOM.scope.classList.remove("tagify--focus");
      });

      const { thumbnail, slide } = await exportSelectedSlideAsBase64();

      const thumbnailImg = document.querySelector("#add-slide-thumbnail");
      thumbnailImg.src = `data:image/png;base64,${thumbnail}`;
      thumbnailImg.dataset.slide = slide;
      console.log("썸네일 이미지 설정 완료");
    }
  });
}

/**
 * 수정 페이지 이벤트 핸들러 등록
 */
function registerEditPageHandlers() {
  console.log("수정 페이지 이벤트 핸들러 등록");

  tryCatch(async () => {
    const slideCache = await getSlideCache();
    console.log("슬라이드 데이터 읽기 성공:", slideCache);

    const editSlideTitleInput = document.querySelector("input[id=edit-slide-title]");
    if (editSlideTitleInput) {
      editSlideTitleInput.value = slideCache.title;
    }

    const thumbnailImg = document.querySelector("#edit-slide-thumbnail");
    if (thumbnailImg && slideCache.thumbnail) {
      thumbnailImg.src = `data:image/png;base64,${slideCache.thumbnail}`;
      console.log("썸네일 이미지 설정 완료");
    }

    const editTagsInput = document.querySelector("input[name=edit-tags]");
    if (editTagsInput) {
      editTagsInput.value = slideCache.tags ? slideCache.tags.join(",") : "";
    }

    const tags = await getTags();

    if (editTagsInput) {
      const editTagify = new Tagify(editTagsInput, {
        whitelist: tags,
        dropdown: {
          maxItems: 5,
          classname: "tags-look",
          enabled: 0,
          closeOnSelect: false,
        },
        maxTags: 10,
      });

      editTagify.DOM.input.addEventListener("focus", function () {
        editTagify.DOM.scope.classList.add("tagify--focus");
      });

      editTagify.DOM.input.addEventListener("blur", function () {
        editTagify.DOM.scope.classList.remove("tagify--focus");
      });

      console.log("수정 페이지 태그 입력 필드에 Tagify 적용됨");
    }
  });
}

/**
 * 계정 페이지 이벤트 핸들러 등록
 */
async function registerAccountPageHandlers() {
  console.log("계정 페이지 이벤트 핸들러 등록");
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
 * 도움말 페이지 (메인) 이벤트 핸들러 등록
 */
function registerHelpPageMainHandlers() {
  console.log("도움말 페이지 (메인) 이벤트 핸들러 등록");
}

/**
 * 도움말 페이지 (설정) 이벤트 핸들러 등록
 */
function registerHelpPageSettingHandlers() {
  console.log("도움말 페이지 (설정) 이벤트 핸들러 등록");
}

/**
 * 업데이트 페이지 이벤트 핸들러 등록
 */
function registerVersionPageHandlers() {
  console.log("버전 페이지 이벤트 핸들러 등록");

  tryCatch(async () => {
    const updateNotification = document.getElementById("update-notification");
    const updateResult = await checkForUpdates();

    if (updateResult.update) {
      updateNotification.textContent = "최신 버전이 아닙니다";
      updateNotification.style.color = "#ff0000";
      updateNotification.style.cursor = "pointer";
      updateNotification.style.textDecoration = "underline";

      updateNotification.addEventListener("click", () => {
        if (updateResult.update_url) {
          window.open(updateResult.update_url, "_blank");
        }
      });
    } else {
      // 최신 버전일 때는 알림을 숨김
      updateNotification.style.display = "none";
    }
  });
}

/**
 * 업데이트 페이지 이벤트 핸들러 등록
 */
function registerUpdatesPageHandlers() {
  const updateDetailsList = document.getElementById("update-details-list");
  const updateLink = document.getElementById("update-link");

  if (!updateDetailsList || !updateLink) return;

  // 업데이트 확인
  checkForUpdates().then(({ update, update_url, update_description }) => {
    if (update) {
      // 업데이트 내용 표시
      updateDetailsList.innerHTML = update_description.map((detail) => `<li>${detail}</li>`).join("");

      // 업데이트 링크 설정
      updateLink.href = update_url;
      updateLink.addEventListener("click", (e) => {
        e.preventDefault();
        window.open(update_url, "_blank");
      });
    } else {
      // 최신 버전인 경우 메인 페이지로 이동
      window.showPage("list-page");
    }
  });
}

/**
 * 설정 페이지 이벤트 핸들러 등록
 */
function registerSettingsPageHandlers() {
  console.log("설정 페이지 이벤트 핸들러 등록");

  const settingsItems = document.querySelectorAll(".settings-item");
  settingsItems.forEach((item) => {
    if (!item.dataset.hasListener) {
      item.addEventListener("click", async function () {
        const settingType = this.dataset.setting;
        console.log(`설정 항목 클릭: ${settingType}`);

        switch (settingType) {
          case "account":
            showPage("account-page");
            break;
          case "help":
            showPage("help-page-setting");
            break;
          case "version":
            showPage("version-page");
            break;
          case "logout":
            try {
              await signOut();
              showPage("main-page");
              clearSlideListCache();
            } catch (error) {
              console.error("로그아웃 오류:", error);
              setMessage("로그아웃 중 오류가 발생했습니다: " + error.message);
            }
            break;
          default:
            break;
        }
      });
      item.dataset.hasListener = "true";
    }
  });
}

/**
 * 공통 이벤트 핸들러 등록
 */
function registerCommonHandlers() {
  const hamburgerMenu = document.getElementById("hamburger-menu");
  if (hamburgerMenu && !hamburgerMenu.dataset.hasListener) {
    hamburgerMenu.addEventListener("click", function () {
      showPage("settings-page");
    });
    hamburgerMenu.dataset.hasListener = "true";
  }

  const settingsBackButton = document.getElementById("settings-page-back");
  if (settingsBackButton && !settingsBackButton.dataset.hasListener) {
    settingsBackButton.addEventListener("click", function () {
      showPage("list-page");
    });
    settingsBackButton.dataset.hasListener = "true";
  }

  const accountBackButton = document.getElementById("account-page-back");
  if (accountBackButton && !accountBackButton.dataset.hasListener) {
    accountBackButton.addEventListener("click", function () {
      showPage("settings-page");
    });
    accountBackButton.dataset.hasListener = "true";
  }

  const helpPageMainBackButton = document.getElementById("help-page-main-back");
  if (helpPageMainBackButton && !helpPageMainBackButton.dataset.hasListener) {
    helpPageMainBackButton.addEventListener("click", function () {
      showPage("main-page");
    });
    helpPageMainBackButton.dataset.hasListener = "true";
  }

  const helpPageSettingBackButton = document.getElementById("help-page-setting-back");
  if (helpPageSettingBackButton && !helpPageSettingBackButton.dataset.hasListener) {
    helpPageSettingBackButton.addEventListener("click", function () {
      showPage("settings-page");
    });
    helpPageSettingBackButton.dataset.hasListener = "true";
  }

  const versionBackButton = document.getElementById("version-page-back");
  if (versionBackButton && !versionBackButton.dataset.hasListener) {
    versionBackButton.addEventListener("click", function () {
      showPage("settings-page");
    });
    versionBackButton.dataset.hasListener = "true";
  }

  const updatesBackButton = document.getElementById("updates-page-back");
  if (updatesBackButton && !updatesBackButton.dataset.hasListener) {
    updatesBackButton.addEventListener("click", function () {
      showPage("settings-page");
    });
    updatesBackButton.dataset.hasListener = "true";
  }
}

// 마지막에 한 번만 export
export {
  showPage,
  registerPageEventHandlers,
  registerMainPageHandlers,
  registerListPageHandlers,
  registerAddPageHandlers,
  registerEditPageHandlers,
  registerAccountPageHandlers,
  registerHelpPageMainHandlers,
  registerHelpPageSettingHandlers,
  registerVersionPageHandlers,
  registerUpdatesPageHandlers,
  registerSettingsPageHandlers,
  registerCommonHandlers,
};
