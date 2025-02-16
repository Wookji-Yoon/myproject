/* global document, Office, console*/

import {
  signIn,
  createFolder,
  createPowerPointFile,
  createJsonFile,
  readJsonFile,
  updateJsonFile,
} from "./graphService";
import { addSlideTag, createJsonData, exportSelectedSlideAsBase64, insertAfterSelectedSlide } from "./functions";

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint && info.platform === Office.PlatformType.PC) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // 네비게이션 초기화
    initializeNavigation();

    // 기본 페이지 표시
    showPage("main-page");

    // Microsoft Graph 관련 이벤트 핸들러에 tryCatch 적용
    document.getElementById("sign-in").onclick = () =>
      tryCatch(async () => {
        const account = await signIn();
        setMessage(`Successfully signed in as ${account.name}!`);
      });

    document.getElementById("create-folder").onclick = () =>
      tryCatch(async () => {
        await createFolder();
        setMessage("Folder created successfully!");
      });

    document.getElementById("create-powerpoint").onclick = () =>
      tryCatch(async () => {
        await createPowerPointFile();
        setMessage("PowerPoint file created successfully!");
      });

    document.getElementById("create-json").onclick = () =>
      tryCatch(async () => {
        await createJsonFile();
        setMessage("JSON file created successfully!");
      });

    document.getElementById("read-json").onclick = () =>
      tryCatch(async () => {
        const jsonData = await readJsonFile();
        displaySlides(jsonData.slides);
        setMessage("JSON 파일을 성공적으로 읽었습니다!");
      });

    document.getElementById("export-selected-slide").addEventListener("click", async (e) => {
      e.preventDefault();
      try {
        const base64 = await exportSelectedSlideAsBase64();

        setMessage(`Exported Slide Base64: ${base64.substring(0, 50)}...`);
      } catch (error) {
        setMessage(`Error exporting slide: ${error.message}`);
      }
    });

    document.getElementById("insert-after-selected-slide").addEventListener("click", async (e) => {
      e.preventDefault();
      const result = await readJsonFile();
      const slides = result.slides;
      console.log(typeof slides);
      console.log(slides[slides.length - 1]);
      const lastSlideId = slides[slides.length - 1].id;
      await insertAfterSelectedSlide(slides, lastSlideId);
      setMessage("Inserted after selected slide!");
    });

    document.getElementById("tag-form").addEventListener("submit", async (e) => {
      e.preventDefault();
      const key = "TOPIC";
      const value = document.getElementById("tag-value").value;
      const userTags = {
        [key]: value,
      };
      try {
        const result = await exportSelectedSlideAsBase64(userTags);
        console.log(result);
        const jsonData = createJsonData(result);
        console.log(jsonData);
        await updateJsonFile(jsonData);
        setMessage(`Exported Slide Base64: ${result.slide.substring(0, 50)}...
        Thumbnail Base64: ${result.thumbnail.substring(0, 50)}...
        Tags: ${JSON.stringify(result.tags)}`);
      } catch (error) {
        setMessage(`Error exporting slide: ${error.message}`);
      }
    });
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
}

// setMessage 함수 추가
function setMessage(message) {
  document.getElementById("message").innerText = message;
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

export { setMessage, clearMessage, tryCatch };
