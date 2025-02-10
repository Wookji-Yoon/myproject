/* global console, Blob, fetch */

import * as msal from "@azure/msal-browser";
import * as MicrosoftGraph from "@microsoft/microsoft-graph-client";

const msalConfig = {
  auth: {
    clientId: "b8f7e758-dbc0-404c-886b-72f7fb9a3414", // 기존 클라이언트 ID 그대로 사용
    authority: "https://login.microsoftonline.com/common",
    redirectUri: "https://localhost:3000",
    loginPersistence: true, // 로그인 상태 유지
  },
  system: {
    loggerOptions: {
      loggerCallback: (level, message, containsPii) => {
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

async function createFolder() {
  await initializeMsal();
  const client = await getGraphClient();

  try {
    // 'myapp' 폴더 생성
    await client.api("/me/drive/root/children").post({
      name: "myapp",
      folder: {},
    });
    console.log("Folder 'myapp' created successfully");
  } catch (error) {
    console.error("Error creating folder:", error);
  }
}

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

async function createJsonFile() {
  await initializeMsal();
  const client = await getGraphClient();

  try {
    // 'myapp' 폴더에 'presentation.json' 파일 생성
    await client.api("/me/drive/root:/myapp/presentation.json:/content").put(
      new Blob(
        [
          JSON.stringify(
            {
              slides: [
                {
                  id: "1",
                  base64: "SGVsbG8gd29ybGQ=",
                  thumbnail: "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQC...",
                  saved_at: "2025-02-10T12:00:00Z",
                  text_content: "This is the first slide content.",
                  tags: {
                    project: "Marketing Campaign",
                    topic: "Social Media Strategy",
                    subtopic: "Instagram Ads",
                  },
                },
                {
                  id: "2",
                  base64: "U2xpZGUgQmFzZTY0IERhdGE=",
                  thumbnail: "iVBORw0KGgoAAAANSUhEUgAAABAAAAAQC...",
                  saved_at: "2025-02-10T12:30:00Z",
                  text_content: "Key points for the meeting.",
                  tags: {
                    project: "Product Launch",
                    topic: "Event Planning",
                    subtopic: "Venue Selection",
                  },
                },
              ],
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

async function readJsonFile() {
  const client = await getGraphClient();

  try {
    // 파일의 다운로드 URL 먼저 획득
    const fileMetadata = await client
      .api("/me/drive/root:/myapp/presentation.json")
      .select("@microsoft.graph.downloadUrl")
      .get();

    // 다운로드 URL로 직접 fetch
    const response = await fetch(fileMetadata["@microsoft.graph.downloadUrl"]);

    if (!response.ok) {
      throw new Error("Network response was not ok");
    }

    const jsonData = await response.json();
    console.log("presentation.json contents:", jsonData);
    return jsonData;
  } catch (error) {
    console.error("Error reading JSON file:", error);
    throw error;
  }
}

export { signIn, createFolder, createPowerPointFile, createJsonFile, readJsonFile };
