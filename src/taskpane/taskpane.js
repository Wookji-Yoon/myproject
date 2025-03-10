/* global document, Office */

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { deleteOneSlideJsonFile, isUserLoggedIn } from "./graphService";
import { handleSignIn, handleExportSlide, showPage, handleInsertSlide, clearSlidesCache } from "./functions";
import { tryCatch, setMessage } from "./utils";

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint && info.platform === Office.PlatformType.PC) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // 네비게이션 초기화
    initializeNavigation();

    // 캐시 초기화 버튼 클릭 처리
    document.getElementById("clear-cache").onclick = () => {
      tryCatch(async () => {
        await clearSlidesCache();
      });
    };

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

        // 클릭한 요소가 삭제 버튼인지 확인(버튼이 아닌 icon인 경우 상위 버튼 요소 찾기)
        if (event.target.classList.contains("delete-icon") || event.target.closest(".delete-icon")) {
          // 클릭된 요소가 아이콘인 경우 상위 버튼 요소 찾기
          const deleteButton = event.target.classList.contains("delete-icon")
            ? event.target
            : event.target.closest(".delete-icon");
          const slideId = deleteButton.dataset.slideId;

          // UI에서 해당 슬라이드 요소 저장 및 제거 (낙관적 UI 업데이트)
          const slideElement = deleteButton.closest(".slide-item");
          const slideContainer = slideElement.parentNode;
          const slideIndex = Array.from(slideContainer.children).indexOf(slideElement);
          const slideClone = slideElement.cloneNode(true); // 복원을 위해 복제

          try {
            // UI에서 먼저 요소 제거 (사용자에게 즉각적인 피드백 제공)
            if (slideElement) {
              slideElement.remove();
              setMessage(`슬라이드 삭제 중...`);
            }

            // 백엔드에서 슬라이드 삭제 시도
            await deleteOneSlideJsonFile(slideId);

            // 성공 시 최종 메시지 표시
            setMessage(`슬라이드 ID: ${slideId}가 삭제되었습니다`);
          } catch (error) {
            // 오류 로깅 (console 경고 대신 Office.context.mailbox.item.notificationMessages 사용)
            // console.error("슬라이드 삭제 실패:", error);

            // 실패 시 UI 복원
            if (slideIndex >= 0) {
              if (slideIndex === 0 && slideContainer.children.length === 0) {
                // 컨테이너가 비어있는 경우 첫 번째 요소로 추가
                slideContainer.appendChild(slideClone);
              } else if (slideIndex >= slideContainer.children.length) {
                // 마지막 요소였던 경우 마지막에 추가
                slideContainer.appendChild(slideClone);
              } else {
                // 중간에 있었던 경우 해당 위치에 삽입
                slideContainer.insertBefore(slideClone, slideContainer.children[slideIndex]);
              }
            }

            setMessage(`슬라이드 삭제에 실패했습니다: ${error.message}`);
          }
        }

        // 수정정 버튼 클릭 처리
        if (event.target.classList.contains("edit-icon") || event.target.closest(".edit-icon")) {
          // 클릭된 요소가 아이콘인 경우 상위 버튼 요소 찾기
          const infoButton = event.target.classList.contains("edit-icon")
            ? event.target
            : event.target.closest(".edit-icon");
          const slideId = infoButton.dataset.slideId;

          setMessage(`수정 버튼 클릭됨. 슬라이드 ID: ${slideId}`);
          // 여기에 정보/수정 로직 추가
        }
      });
    });

    // 3. add-page에서 슬라이드 추가하기 버튼 클릭시 export 처리
    document.getElementById("add-slide-button").onclick = () =>
      tryCatch(async () => {
        await handleExportSlide();
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
