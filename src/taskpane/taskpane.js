/* global document, Office */

/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import { isUserLoggedIn } from "./graphService";
import { handleSignIn, handleExportSlide, showPage, insertAfterSelectedSlide } from "./functions";
import { tryCatch, setMessage } from "./utils";
import { readJsonFile } from "./graphService";

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

          // 메시지 표시
          setMessage(`슬라이드 ${slideId}를 선택했습니다`);

          const jsonData = await readJsonFile(`/me/drive/root:/myapp/slides.json`);
          await insertAfterSelectedSlide(jsonData.slides, slideId);

          // 여기에 슬라이드 삽입 로직을 추가할 수 있습니다
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
