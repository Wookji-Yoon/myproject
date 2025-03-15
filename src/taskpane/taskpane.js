/* global document, Office, window, setTimeout, console */

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { isUserLoggedIn, readJsonFile } from "./graphService";
import {
  handleSignIn,
  handleExportSlide,
  showPage,
  handleInsertSlide,
  handleEditSlide,
  handleDeleteIconClick,
  handleEditIconClick,
  handleTitleSearch,
  handleTagSearch,
} from "./functions";
import { tryCatch } from "./utils";

// Make showPage function globally accessible
window.showPage = showPage;

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint && info.platform === Office.PlatformType.PC) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // 네비게이션 초기화
    initializeNavigation();

    // 로그인 상태에 따라 기본 페이지 표시
    tryCatch(async () => {
      const isLoggedIn = await isUserLoggedIn();
      if (isLoggedIn) {
        showPage("add-page");
      } else {
        showPage("main-page");
      }
    });

    // 1. main-page에서 sign-in 버튼 클릭시 로그인 처리
    document.getElementById("sign-in").onclick = () =>
      tryCatch(async () => {
        await handleSignIn();
      });

    // 2. list-page에서 슬라이드 클릭시 슬라이드 삽입
    document.getElementById("slides-container").addEventListener("click", (event) => {
      tryCatch(async () => {
        // 클릭된 요소가 썸네일 이미지인지 확인
        if (event.target.classList.contains("slide-thumbnail")) {
          // 해당 슬라이드의 ID 가져오기 (alt 텍스트에서 추출)
          const altText = event.target.alt;
          const slideId = altText.replace("슬라이드 썸네일", "");

          // 핸들러 함수 호출
          await handleInsertSlide(slideId);
        }

        // 클릭한 요소가 삭제 버튼인지 확인
        if (event.target.classList.contains("delete-icon") || event.target.closest(".delete-icon")) {
          await handleDeleteIconClick(event);
        }

        // 수정 버튼 클릭 처리
        if (event.target.classList.contains("edit-icon") || event.target.closest(".edit-icon")) {
          await handleEditIconClick(event);
        }
      });
    });

    // 슬라이드 추가 버튼 클릭 이벤트
    document.getElementById("add-new-slide-button").addEventListener("click", () => {
      showPage("add-page");
    });

    // list page에서 serach filter 클릭시 처리
    // 필터 탭 클릭 이벤트
    document.querySelectorAll(".filter-tab").forEach((tab) => {
      tab.onclick = () => {
        // 이미 활성화된 탭을 클릭한 경우 아무 작업도 하지 않음
        if (tab.classList.contains("active")) return;

        // 기존 활성화된 탭에서 active 클래스 제거
        document.querySelector(".filter-tab.active").classList.remove("active");
        // 클릭한 탭에 active 클래스 추가
        tab.classList.add("active");

        if (tab.getAttribute("data-filter") === "title") {
          document.getElementById("tag-search-input").value = "";
          document.getElementById("tag-search-wrapper").classList.add("hidden");
          document.getElementById("title-search-wrapper").classList.remove("hidden");
        } else {
          document.getElementById("title-search-input").value = "";
          document.getElementById("title-search-wrapper").classList.add("hidden");
          document.getElementById("tag-search-wrapper").classList.remove("hidden");
        }
      };
    });

    // 검색 버튼 클릭 이벤트
    document.getElementById("title-search-button").onclick = () =>
      tryCatch(async () => {
        const searchInput = document.getElementById("title-search-input");
        const searchButton = document.getElementById("title-search-button");

        //UI 처리
        searchInput.blur();
        searchInput.disabled = true;
        searchButton.innerHTML = '<i class="ms-Icon ms-Icon--ProgressRingDots" aria-hidden="true"></i>';

        //0.5초뒤 복구
        setTimeout(() => {
          searchInput.disabled = false;
          searchButton.innerHTML = '<i class="ms-Icon ms-Icon--Search" aria-hidden="true"></i>';
        }, 500);

        await handleTitleSearch(searchInput.value);
      });

    // 태그 검색 버튼 클릭 이벤트
    document.getElementById("tag-search-button").onclick = () =>
      tryCatch(async () => {
        const searchInput = document.getElementById("tag-search-input");
        await handleTagSearch(searchInput.value);
      });

    document.getElementById("title-search-input").addEventListener("keyup", (event) => {
      if (event.key === "Enter") {
        document.getElementById("title-search-button").click();
      }
    });

    // 3. add-page에서 슬라이드 추가하기 버튼 클릭시 export 처리
    document.getElementById("add-slide-button").onclick = () =>
      tryCatch(async () => {
        await handleExportSlide();
      });

    // 4. edit-page에서 슬라이드 수정하기 버튼 클릭시 수정 처리
    document.getElementById("edit-slide-button").onclick = () =>
      tryCatch(async () => {
        await handleEditSlide();
      });
  }
});

function initializeNavigation() {
  const navButtons = document.querySelectorAll(".nav-button");
  navButtons.forEach((button) => {
    button.addEventListener("click", function () {
      const pageId = this.getAttribute("data-page");
      showPage(pageId);
    });
  });
}
