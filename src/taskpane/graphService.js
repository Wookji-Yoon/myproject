/* global console, Blob, fetch */

import * as msal from "@azure/msal-browser";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-client";

import { subtractArrays } from "./utils";
import { sampleDict } from "./sample";

const msalConfig = {
  auth: {
    clientId: "b8f7e758-dbc0-404c-886b-72f7fb9a3414", // 기존 클라이언트 ID 그대로 사용
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "https://localhost:3000",
    loginPersistence: true, // 로그인 상태 유지
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message) => {
        console.log(`MSAL Log: ${message}`);
      },
      piiLoggingEnabled: false,
    },
  },
  cache: {
    cacheLocation: "localStorage", // 브라우저 로컬 스토리지에 캐시
    storeAuthStateInCookie: true, // 쿠키에 인증 상태 저장
  },
};

const msalInstance = new msal.PublicClientApplication(msalConfig);

// MSAL 초기화 함수 추가
async function initializeMsal() {
  try {
    // MSAL v3에서는 initialize() 메서드만 사용
    await msalInstance.initialize();
  } catch (error) {
    console.error("MSAL initialization error:", error);
    throw error;
  }
}

// 로그인 요청 파라미터
const loginRequest = {
  scopes: ["Files.ReadWrite", "User.Read"],
};

async function signIn() {
  try {
    // MSAL 초기화 확인 및 실행
    await initializeMsal();

    // 로그인 시도
    const response = await msalInstance.loginPopup(loginRequest);
    msalInstance.setActiveAccount(response.account);
    return response.account;
  } catch (error) {
    console.error("Login failed:", error);
    throw error;
  }
}

async function getGraphClient() {
  // MSAL 초기화 추가
  await initializeMsal();

  // 액세스 토큰 획득
  const account = msalInstance.getActiveAccount();
  if (!account) {
    throw new Error("No active account");
  }

  const request = {
    scopes: loginRequest.scopes,
    account: account,
  };

  const response = await msalInstance.acquireTokenSilent(request);

  // Microsoft Graph 클라이언트 생성
  return MicrosoftGraph.Client.init({
    authProvider: (done) => {
      done(null, response.accessToken);
    },
  });
}

async function createFolder(folderName = "myapp") {
  await initializeMsal();
  const client = await getGraphClient();

  try {
    // 폴더 존재 여부 확인
    try {
      await client.api(`/me/drive/root:/${folderName}`).get();
      console.log(`Folder '${folderName}' already exists`);
      return true; // 폴더가 이미 존재함
    } catch (error) {
      // 폴더가 없으면 생성
      if (error.statusCode === 404) {
        await client.api("/me/drive/root/children").post({
          name: folderName,
          folder: {},
        });
        console.log(`Folder '${folderName}' created successfully`);
        return true;
      } else {
        throw error;
      }
    }
  } catch (error) {
    console.error("Error creating folder:", error);
    return false;
  }
}

/* 현재 사용되지 않는 함수
async function createPowerPointFile() {
  await initializeMsal();
  const client = await getGraphClient();

  try {
    // 'myapp' 폴더에 'mypowerpoint.pptx' 파일 생성
    await client
      .api("/me/drive/root:/myapp/mypowerpoint.pptx:/content")
      .put(new Blob([], { type: "application/vnd.openxmlformats-officedocument.presentationml.presentation" }));

    console.log("PowerPoint file created successfully");
  } catch (error) {
    console.error("Error creating PowerPoint file:", error);
  }
}
*/

async function fileExists(filePath) {
  await initializeMsal();
  const client = await getGraphClient();

  try {
    // 파일 존재 여부는 메타데이터로 확인하므로 /content를 추가하지 않음
    await client.api(filePath).get();
    return true;
  } catch (error) {
    if (error.statusCode === 404) {
      return false;
    }
    throw error;
  }
}

async function createJsonFile(filePath = "/me/drive/root:/myapp/") {
  await initializeMsal();
  const client = await getGraphClient();

  try {
    // 기본 JSON 구조를 exportSelectedSlideAsBase64의 출력 형식과 일치시킴
    await client.api(filePath + "slides.json:/content").put(
      new Blob(
        [
          JSON.stringify(
            {
              slides: [sampleDict],
            },
            null,
            2
          ),
        ],
        { type: "application/json" }
      )
    );

    // Tag 저장용 Json을 따로 만듦
    await client.api(filePath + "tags.json:/content").put(
      new Blob(
        [
          JSON.stringify(
            {
              tags: ["샘플", "sample"],
            },
            null,
            2
          ),
        ],
        { type: "application/json" }
      )
    );

    console.log("presentation.json created successfully");
  } catch (error) {
    console.error("Error creating JSON file:", error);
  }
}

async function readJsonFile(path = "/me/drive/root:/myapp/slides.json") {
  try {
    console.log("readJsonFile 함수 시작");
    const client = await getGraphClient();
    console.log("Graph 클라이언트 준비 완료");

    try {
      // 파일의 다운로드 URL 먼저 획득
      console.log("파일 다운로드 URL 요청 중...");
      const fileMetadata = await client.api(path).select("@microsoft.graph.downloadUrl").get();

      console.log("파일 다운로드 URL 획득 성공");

      // 다운로드 URL로 직접 fetch
      console.log("파일 콘텐츠 가져오는 중...");
      const response = await fetch(fileMetadata["@microsoft.graph.downloadUrl"]);

      if (!response.ok) {
        throw new Error(`Network response was not ok: ${response.status} ${response.statusText}`);
      }

      const jsonData = await response.json();
      console.log("JSON 파일 내용:", jsonData);
      return jsonData;
    } catch (error) {
      console.error("Error reading JSON file:", error);
      throw error;
    }
  } catch (error) {
    console.error("Error in readJsonFile function:", error);
    throw error;
  }
}

async function updateJsonFile(jsonData) {
  const client = await getGraphClient();

  try {
    // 기존존 JSON 파일 읽기
    const existingData = await readJsonFile();

    // 새 슬라이드를 기존 슬라이드 목록의 가장 앞에 추가
    existingData.slides.unshift(jsonData);

    // Microsoft Graph API를 사용하여 파일 업데이트
    await client.api("/me/drive/root:/myapp/slides.json:/content").put(JSON.stringify(existingData, null, 2));

    console.log("JSON file updated successfully");
  } catch (error) {
    console.error("Error updating JSON file:", error);
    throw error;
  }

  try {
    const tagJsonData = await readJsonFile("/me/drive/root:/myapp/tags.json");
    console.log("tagJsonData", tagJsonData);
    console.log("jsonData", jsonData);

    // 두 태그 배열을 합치기 (중복 허용용)
    const combinedTags = [...tagJsonData.tags, ...jsonData.tags];
    console.log("combinedTags", combinedTags);

    //combinedTags를 가나다순으로 정렬
    const sortTags = combinedTags.sort();

    tagJsonData.tags = sortTags;

    // 업데이트된 태그 저장
    await client.api("/me/drive/root:/myapp/tags.json:/content").put(JSON.stringify(tagJsonData, null, 2));
    console.log("Tags updated successfully");
  } catch (error) {
    console.error("Error updating tags:", error);
    throw error;
  }
}

async function deleteOneSlideJsonFile(slideId) {
  const client = await getGraphClient();
  const existingData = await readJsonFile();

  try {
    //삭제할 슬라이드 찾아서 제거
    const updatedData = existingData.slides.filter((slide) => slide.id !== slideId);
    const updatedJsonData = {
      slides: updatedData,
    };
    await client.api("/me/drive/root:/myapp/slides.json:/content").put(JSON.stringify(updatedJsonData, null, 2));
  } catch (error) {
    console.error("Error deleting slide:", error);
    throw error;
  }

  try {
    // 태그목록에서 제거
    const tagJsonData = await readJsonFile("/me/drive/root:/myapp/tags.json");

    const targetSlide = existingData.slides.find((slide) => slide.id === slideId);
    const deletedTags = targetSlide.tags;
    const updatedTags = subtractArrays(tagJsonData.tags, deletedTags);
    const updatedTagsJsonData = {
      tags: updatedTags,
    };

    await client.api("/me/drive/root:/myapp/tags.json:/content").put(JSON.stringify(updatedTagsJsonData, null, 2));
  } catch (error) {
    console.error("Error deleting tags:", error);
    throw error;
  }
}

async function editJsonFile(updatedSlide) {
  const client = await getGraphClient();

  try {
    const existingData = await readJsonFile();
    existingData.slides.find((slide) => slide.id === updatedSlide.id).title = updatedSlide.title;
    existingData.slides.find((slide) => slide.id === updatedSlide.id).tags = updatedSlide.tags;
    await client.api("/me/drive/root:/myapp/slides.json:/content").put(JSON.stringify(existingData, null, 2));
  } catch (error) {
    console.error("Error editing JSON file:", error);
    throw error;
  }

  try {
    const tagJsonData = await readJsonFile("/me/drive/root:/myapp/tags.json");
    const combinedTags = [...tagJsonData.tags, ...updatedSlide.tags];
    tagJsonData.tags = combinedTags;
    await client.api("/me/drive/root:/myapp/tags.json:/content").put(JSON.stringify(tagJsonData, null, 2));
  } catch (error) {
    console.error("Error editing tags:", error);
    throw error;
  }
}

async function isUserLoggedIn() {
  try {
    await initializeMsal();
    const account = msalInstance.getActiveAccount();
    return !!account; // Returns true if account exists, false otherwise
  } catch (error) {
    console.error("Error checking login status:", error);
    return false;
  }
}

async function getAccountInfo() {
  try {
    await initializeMsal();
    const account = msalInstance.getActiveAccount();

    if (!account) {
      throw new Error("사용자가 로그인되어 있지 않습니다.");
    }

    // Graph API를 통해 더 자세한 사용자 정보 가져오기
    const client = await getGraphClient();
    const userDetails = await client.api("/me").get();

    return {
      basicInfo: account, // MSAL에서 제공하는 기본 계정 정보
      detailedInfo: userDetails, // Graph API에서 제공하는 상세 정보
    };
  } catch (error) {
    console.error("계정 정보 가져오기 실패:", error);
    throw error;
  }
}

async function signOut() {
  try {
    await initializeMsal();

    // 모든 계정에서 로그아웃 (single account 모드에서도 안전)
    const logoutRequest = {
      account: msalInstance.getActiveAccount(),
      postLogoutRedirectUri: msalConfig.auth.redirectUri,
    };

    await msalInstance.logoutPopup(logoutRequest);
    console.log("로그아웃 성공");
    return true;
  } catch (error) {
    console.error("로그아웃 실패:", error);
    throw error;
  }
}

export {
  initializeMsal,
  signIn,
  fileExists,
  createJsonFile,
  readJsonFile,
  updateJsonFile,
  createFolder,
  deleteOneSlideJsonFile,
  editJsonFile,
  isUserLoggedIn,
  getAccountInfo,
  signOut,
};
