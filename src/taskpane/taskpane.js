/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, PowerPoint, console, OfficeExtension, FileReader*/

// TODO1: Import Base64-encoded string for image.
import { base64Image } from "../../base64Image";

// onReady function gets called once the Office.js library has initialized.
Office.onReady((info) => {
  if (info.host === Office.HostType.PowerPoint && info.platform === Office.PlatformType.PC) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    // TODO2: Assign event handler for insert-image button.
    document.getElementById("insert-image").onclick = () => clearMessage(insertImage);

    // TODO4: Assign event handler for insert-text button.
    document.getElementById("insert-text").onclick = () => clearMessage(insertText);
    // TODO6: Assign event handler for get-slide-metadata button.
    document.getElementById("get-slide-metadata").onclick = () => clearMessage(getSlideMetadata);
    // TODO8: Assign event handlers for add-slides and the four navigation buttons.
    document.getElementById("add-slides").onclick = () => tryCatch(addSlides);
    document.getElementById("go-to-first-slide").onclick = () => clearMessage(goToFirstSlide);
    document.getElementById("go-to-next-slide").onclick = () => clearMessage(goToNextSlide);
    document.getElementById("go-to-previous-slide").onclick = () => clearMessage(goToPreviousSlide);
    document.getElementById("go-to-last-slide").onclick = () => clearMessage(goToLastSlide);
    document.getElementById("get-slide-index").onclick = () => clearMessage(getSelectedSlideIndex);
    document.getElementById("add").onclick = () => addSlideWithMatchingLayout();
    document.getElementById("delete").onclick = () => deleteSlide();
    document.getElementById("file").onchange = () => storeFileAsBase64();
    document.getElementById("insert-all").onclick = () => insertAllSlides();
    document.getElementById("insert-target").onclick = () => insertAfterSelectedSlide();
  }
});

// TODO3: Define the insertImage function.
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

// TODO5: Define the insertText function.
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

// TODO7: Define the getSlideMetadata function.
function getSlideMetadata() {
  Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Failed) {
      setMessage("Error: " + asyncResult.error.message);
    } else {
      setMessage("Slide metadata: " + JSON.stringify(asyncResult.value));
    }
  });
}

// TODO9: Define the addSlides and navigation functions.
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

async function clearMessage(callback) {
  document.getElementById("message").innerText = "";
  await callback();
}

function setMessage(message) {
  document.getElementById("message").innerText = message;
}

// Default helper for invoking an action and handling errors.
async function tryCatch(callback) {
  try {
    document.getElementById("message").innerText = "";
    await callback();
  } catch (error) {
    setMessage("Error: " + error.toString());
  }
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

async function addMultipleSlideTags() {
  await PowerPoint.run(async function (context) {
    const slide = context.presentation.slides.getItemAt(0);
    slide.tags.add("OCEAN", "Arctic");
    slide.tags.add("PLANET", "Jupiter");

    await context.sync();
  });
}

async function updateTag() {
  await PowerPoint.run(async function (context) {
    const slide = context.presentation.slides.getItemAt(0);
    slide.tags.add("PLANET", "Mars");

    await context.sync();
  });
}

async function deleteTag() {
  await PowerPoint.run(async function (context) {
    const slide = context.presentation.slides.getItemAt(0);
    slide.tags.delete("PLANET");

    await context.sync();
  });
}

async function addTagToSelectedSlide() {
  await PowerPoint.run(async function (context) {
    let selectedSlideIndex = await getSelectedSlideIndex();
    selectedSlideIndex = selectedSlideIndex - 1;
    const slide = context.presentation.slides.getItemAt(selectedSlideIndex);
    slide.tags.add("CUSTOMER_TYPE", "Premium");

    await context.sync();
  });
}

async function deleteSlidesByAudience() {
  await PowerPoint.run(async function (context) {
    const slides = context.presentation.slides;
    slides.load("tags/key, tags/value");

    await context.sync();

    for (let i = 0; i < slides.items.length; i++) {
      let currentSlide = slides.items[i];
      for (let j = 0; j < currentSlide.tags.items.length; j++) {
        let currentTag = currentSlide.tags.items[j];
        if (currentTag.key === "CUSTOMER_TYPE" && currentTag.value === "Premium") {
          currentSlide.delete();
        }
      }
    }

    await context.sync();
  });
}
