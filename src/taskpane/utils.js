/* global Office, OfficeExtension, console */

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

export { getSelectedSlideIndex, getSelectedSlideId };
