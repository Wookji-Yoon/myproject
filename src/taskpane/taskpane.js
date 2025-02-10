/* global document, Office */

import { signIn, createFolder, createPowerPointFile, createJsonFile, readJsonFile } from "./graphService";
import { addSlideTag, exportSelectedSlideAsBase64, insertAfterSelectedSlide } from "./functions";

Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint && info.platform === Office.PlatformType.PC) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

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
        setMessage(`JSON file read successfully! Data: ${JSON.stringify(jsonData)}`);
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
      await insertAfterSelectedSlide();
      setMessage("Inserted after selected slide!");
    });

    document.getElementById("tag-form").addEventListener("submit", async (e) => {
      e.preventDefault();
      const key = "TOPIC";
      const value = document.getElementById("tag-value").value;
      await addSlideTag(key, value);
      setMessage(`태그 "${value}" 추가 완료!`);
    });
  }
});

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

export { setMessage, clearMessage, tryCatch };
