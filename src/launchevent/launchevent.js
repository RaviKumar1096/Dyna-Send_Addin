function onMessageComposeHandler(event) {
  setSubject(event);
}

async function onMessageSendHandler(event) {
  OnsenSetSignature(event);
}

async function OnsenSetSignature(event) {
  Office.context.mailbox.item.subject.getAsync({ asyncContext: event }, function (asyncResult) {
    composeSignatureonSend(asyncResult.asyncContext, event).then(() => {
      asyncResult.asyncContext.completed({ allowEvent: true });
    });
  });
}

function composeSignatureonSend(event, event1) {
  return Office.context.mailbox.item.getComposeTypeAsync({ asyncContext: event }, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      var _settings = Office.context.roamingSettings;
      var property = _settings.get("SetFlag");
      _settings.set("SetFlag", false);
      Office.context.roamingSettings.saveAsync(function (result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          console.error(`Action failed with message ${result.error.message}`);
        } else {
          console.log(`Settings saved with status: ${result.status}`);
        }
      });
      const sign = extractSignature2(asyncResult.value.composeType, property);
      setSignatureonSend(event1, sign);
    } else {
      console.error(asyncResult.error);
    }
  });
}

function setSignatureonSend(event1, sign1) {
  return Office.context.mailbox.item.body.setSignatureAsync(
    sign1,
    {
      coercionType: Office.CoercionType.Html,
    },
    function (asyncResult) {
      console.log("set_signature - " + JSON.stringify(asyncResult));
      event1.completed({ allowEvent: true });
    }
  );
}

async function setSubject(event) {
  var _settings = Office.context.roamingSettings;
  var property = _settings.get("SetFlag");
  _settings.set("SetFlag", false);
  Office.context.roamingSettings.saveAsync(function (result) {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      console.error(`Action failed with message ${result.error.message}`);
    } else {
      console.log(`Settings saved with status: ${result.status}`);
    }
  });
  return getSignature().then(async (data) => await composeSignature(data.data, event, property));
}

function saveMyAppSettingsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  } else {
  }
}

function getSignature() {
  var requestOptions = {
    method: "GET",
    redirect: "follow",
  };

  return fetch("https://manage.dynasend.net/s/outlook-bundle?email=ravi@metadesignsolutions.co.uk", requestOptions)
    .then((response) => response.json())
    .catch((error) => console.log("error", error));
}

async function composeSignature(data, event, property) {
  await Office.context.mailbox.item.getComposeTypeAsync({ asyncContext: event }, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      const sign = extractSignature(data, asyncResult.value.composeType, property);
      setSignature(sign);
    } else {
      console.error(asyncResult.error);
    }
  });
}

function extractSignature2(composeType, isChecked) {
  var requestOptions = {
    method: "GET",
    redirect: "follow",
  };

  fetch("https://manage.dynasend.net/s/outlook-bundle?email=ravi@metadesignsolutions.co.uk", requestOptions)
    .then((response) => {
      var signature = response.json();
      const internalSignatue = signature.filter((x) => x.name == "Internal");
      const defaultSignature = signature.filter((x) => x.name == "Default");
      const replySignature = signature.filter((x) => x.name == "Reply");

      if (composeType == "newMail" || isChecked == false) {
        return defaultSignature[0].html_content;
      } else if ((composeType == "reply" || composeType == "forward") && isChecked == true) {
        return replySignature[0].html_content;
      } else {
        return internalSignatue[0].html_content;
      }
    })
    .catch((error) => console.log("error", error));
}

function extractSignature(signature, composeType, isChecked) {
  const internalSignatue = signature.filter((x) => x.name == "Internal");
  const defaultSignature = signature.filter((x) => x.name == "Default");
  const replySignature = signature.filter((x) => x.name == "Reply");
  if (composeType == "newMail" || isChecked == false) {
    return defaultSignature[0].html_content;
  } else if ((composeType == "reply" || composeType == "forward") && isChecked == true) {
    return replySignature[0].html_content;
  } else {
    return internalSignatue[0].html_content;
  }
}

function setSignature(signType) {
  Office.context.mailbox.item.body.setSignatureAsync(
    signType,
    {
      coercionType: Office.CoercionType.Html,
    },
    function (asyncResult1) {
      console.log("set_signature - " + JSON.stringify(asyncResult1));
    }
  );
}


// 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
