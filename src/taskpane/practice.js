/* global document, Office, PowerPoint, console, OfficeExtension, FileReader */

import { base64Image } from "../../base64Image";
import { setMessage } from "./taskpane";

// Slide 관련 함수들
function insertImage() {
  // Call Office.js to insert the image into the document.
  Office.context.document.setSelectedDataAsync(
    base64Image,
    {
      coercionType: Office.CoercionType.Image,
    },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        setMessage("Error: " + asyncResult.error.message);
      }
    }
  );
}

function insertText() {
  Office.context.document.setSelectedDataAsync(
    "Hello World!",
    {
      coercionType: Office.CoercionType.Text,
    },
    (asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Failed) {
        setMessage("Error: " + asyncResult.error.message);
      }
    }
  );
}

function getSlideMetadata() {
  Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      setMessage("Error: " + asyncResult.error.message);
    } else {
      setMessage("Slide metadata: " + JSON.stringify(asyncResult.value));
    }
  });
}

async function addSlides() {
  await PowerPoint.run(async function (context) {
    context.presentation.slides.add();
    context.presentation.slides.add();

    await context.sync();

    goToLastSlide();
    setMessage("Success: Slide added");
  });
}

function goToFirstSlide() {
  Office.context.document.goToByIdAsync(Office.Index.First, Office.GoToType.Index, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      setMessage("Error: " + asyncResult.error.message);
    }
  });
}

function goToLastSlide() {
  Office.context.document.goToByIdAsync(Office.Index.Last, Office.GoToType.Index, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      setMessage("Error: " + asyncResult.error.message);
    }
  });
}

function goToNextSlide() {
  Office.context.document.goToByIdAsync(Office.Index.Next, Office.GoToType.Index, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      setMessage("Error: " + asyncResult.error.message);
    }
  });
}

function goToPreviousSlide() {
  Office.context.document.goToByIdAsync(Office.Index.Previous, Office.GoToType.Index, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      setMessage("Error: " + asyncResult.error.message);
    }
  });
}

function getSelectedSlideIndex() {
  return new OfficeExtension.Promise(function (resolve, reject) {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
      try {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          reject(console.error(asyncResult.error.message));
        } else {
          resolve(asyncResult.value.slides[0].index);
        }
      } catch (error) {
        reject(console.log(error));
      }
    });
  });
}

async function addSlideWithMatchingLayout() {
  await PowerPoint.run(async function (context) {
    let selectedSlideIndex = await getSelectedSlideIndex();

    // Decrement the index because the value returned by getSelectedSlideIndex()
    // is 1-based, but SlideCollection.getItemAt() is 0-based.
    const realSlideIndex = selectedSlideIndex - 1;
    const selectedSlide = context.presentation.slides.getItemAt(realSlideIndex).load("slideMaster/id, layout/id");

    await context.sync();

    context.presentation.slides.add({
      slideMasterId: selectedSlide.slideMaster.id,
      layoutId: selectedSlide.layout.id,
    });

    await context.sync();
  });
}

async function deleteSlide() {
  await PowerPoint.run(async function (context) {
    let selectedSlideIndex = await getSelectedSlideIndex();
    const realSlideIndex = selectedSlideIndex - 1;
    const selectedSlide = context.presentation.slides.getItemAt(realSlideIndex);
    selectedSlide.delete();

    await context.sync();
  });
}

let chosenFileBase64;

function storeFileAsBase64() {
  const reader = new FileReader();

  reader.onload = (event) => {
    const startIndex = reader.result.toString().indexOf("base64,");
    const copyBase64 = reader.result.toString().substring(startIndex + 7);
    chosenFileBase64 = copyBase64;
  };

  const myFile = document.getElementById("file");
  if (myFile.files && myFile.files[0]) {
    reader.readAsDataURL(myFile.files[0]);
  }
}

async function insertAllSlides() {
  console.log("hi");

  await PowerPoint.run(async function (context) {
    context.presentation.insertSlidesFromBase64(chosenFileBase64);
    await context.sync();
  });
}

async function insertAfterSelectedSlide() {
  await PowerPoint.run(async function (context) {
    const selectedSlideId = await getSelectedSlideId();

    context.presentation.insertSlidesFromBase64(chosenFileBase64, {
      formatting: PowerPoint.InsertSlideFormatting.keepSourceFormatting,
      targetSlideId: selectedSlideId + "#",
      sourceSlideIds: ["267#763315295", "256#", "#926310875", "1270#"],
    });

    await context.sync();
  });
}

function getSelectedSlideId() {
  return new OfficeExtension.Promise(function (resolve, reject) {
    Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, function (asyncResult) {
      try {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          reject(console.error(asyncResult.error.message));
        } else {
          resolve(asyncResult.value.slides[0].id);
        }
      } catch (error) {
        reject(console.log(error));
      }
    });
  });
}

export {
  insertImage,
  insertText,
  getSlideMetadata,
  addSlides,
  goToFirstSlide,
  goToLastSlide,
  goToNextSlide,
  goToPreviousSlide,
  getSelectedSlideIndex,
  addSlideWithMatchingLayout,
  deleteSlide,
  storeFileAsBase64,
  insertAllSlides,
  insertAfterSelectedSlide,
};
