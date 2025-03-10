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
import Tagify from "@yaireo/tagify";

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
        showPage("list-page");
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

    // list page에서 serach filter 클릭시 처리
    // 필터 탭 클릭 이벤트
    document.querySelectorAll(".filter-tab").forEach((tab) => {
      tab.onclick = () => {
        document.querySelector(".filter-tab.active").classList.remove("active");
        tab.classList.add("active");
        const activeFilter = document.querySelector(".filter-tab.active").getAttribute("data-filter");
        const searchInputContainer = document.getElementById("search-input").parentElement;
        const originalInput = document.getElementById("search-input");

        // 기존 입력 필드 완전히 제거하고 새로 생성
        if (originalInput) {
          const newInput = document.createElement("input");
          newInput.id = "search-input";
          newInput.type = "text";
          newInput.placeholder = activeFilter === "title" ? "제목으로 검색..." : "태그로 검색...";
          newInput.className = originalInput.className;

          searchInputContainer.removeChild(originalInput);
          searchInputContainer.insertBefore(newInput, searchInputContainer.firstChild);

          // 엔터키 이벤트 다시 등록
          newInput.addEventListener("keypress", (event) => {
            if (event.key === "Enter") {
              document.getElementById("search-button").click();
            }
          });

          if (activeFilter === "tag") {
            tryCatch(async () => {
              const tagJsonData = await readJsonFile("/me/drive/root:/myapp/tags.json");
              console.log("태그 JSON 데이터 읽기 성공:", tagJsonData);
              new Tagify(newInput, {
                whitelist: [...new Set([...tagJsonData.tags])],
                dropdown: {
                  maxItems: 5,
                  classname: "tags-look",
                  enabled: 0,
                  clearOnSelect: false,
                },
                enforceWhitelist: true,
              });
            });
          }
        }
      };
    });

    // list page에서 serach 처리

    // 검색 버튼 클릭 이벤트
    document.getElementById("search-button").onclick = () =>
      tryCatch(async () => {
        const searchInput = document.getElementById("search-input");
        const activeFilter = document.querySelector(".filter-tab.active").getAttribute("data-filter");
        searchInput.blur();
        //searchinput 1초 동안 비활성화
        searchInput.disabled = true;
        setTimeout(() => {
          searchInput.disabled = false;
        }, 1000);
        //1초 동안 뒤에 검색 버튼을 spinner로 변경
        document.getElementById("search-button").innerHTML =
          '<i class="ms-Icon ms-Icon--ProgressRingDots" aria-hidden="true"></i>';
        setTimeout(() => {
          document.getElementById("search-button").innerHTML =
            '<i class="ms-Icon ms-Icon--Search" aria-hidden="true"></i>';
        }, 1000);

        if (activeFilter === "title") {
          await handleTitleSearch(searchInput.value);
          //input에 되어있는 focus 제거
        } else if (activeFilter === "tag") {
          await handleTagSearch(searchInput.value);
        }
      });
    // 검색 입력창 엔터키 이벤트
    document.getElementById("search-input").addEventListener("keypress", (event) => {
      if (event.key === "Enter") {
        document.getElementById("search-button").click();
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
