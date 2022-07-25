import { ABBREVIATED_SIGN, COMPOSE_FORWARD, COMPOSE_NEWMAIL, COMPOSE_REPLY, ErrorMsg, SIGN_DEFAULT, SIGN_FORWARD, SIGN_INTERNAL, SIGN_REPLY } from "../../constants/SignatureConstants";

export function extractSignature (signature, composeType, inDomian,isChecked) {
  const isAbbreviated = localStorage.getItem(ABBREVIATED_SIGN);
  const internalSignatue = signature.filter((x) => x.name == SIGN_INTERNAL);
  const defaultSignature = signature.filter((x) => x.name == SIGN_DEFAULT);
  const replySignature = signature.filter((x) => x.name == SIGN_REPLY);

  if (defaultSignature.length > 0 && internalSignatue.length < 0 && replySignature.length < 0) {
    return defaultSignature[0].html_content;
  } else if (defaultSignature.length > 0 && replySignature.length > 0 && internalSignatue.length < 0) {
    if (composeType == COMPOSE_NEWMAIL) {
      return defaultSignature[0].html_content;
    } else {
      return replySignature[0].html_content;
    }
  } else if (defaultSignature.length > 0 && internalSignatue.length > 0 && replySignature.length > 0) {
    if (composeType == "newMail" || isChecked==false) {
      return defaultSignature[0].html_content;
    } else if ((composeType == "reply" || composeType == "forward") && isChecked==true) {
      return replySignature[0].html_content;
    } else {
      return internalSignatue[0].html_content;
    } 
  } else {
    return ErrorMsg;
  }
};


export async function composeSignatureOnSend(data, inDomian, async) {
  await Office.context.mailbox.item.getComposeTypeAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
      var _settings = Office.context.roamingSettings;
      var property=_settings.get("SetFlag");
      _settings.set("SetFlag",false);
      Office.context.roamingSettings.saveAsync(function(result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          console.error(`Action failed with message ${result.error.message}`);
        } else {
          console.log(`Settings saved with status: ${result.status}`);
        }
      });
      const sign = extractSignature(data, asyncResult.value.composeType, inDomian,property);
      console.log(
        "getComposeTypeAsync succeeded with composeType: " +
          asyncResult.value.composeType +
          " and coercionType: " +
          asyncResult.value.coercionType
      );
      setSignatureOnSend(sign, async);
    } else {
      console.error(asyncResult.error);
    }
  });
}



export async function composeSignature(data, inDomian) {
    await Office.context.mailbox.item.getComposeTypeAsync(function (asyncResult) {
      if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {

    var _settings = Office.context.roamingSettings;
    var property=_settings.get("SetFlag");
    _settings.set("SetFlag",false);
    Office.context.roamingSettings.saveAsync(function(result) {
      if (result.status !== Office.AsyncResultStatus.Succeeded) {
        console.error(`Action failed with message ${result.error.message}`);
      } else {
        console.log(`Settings saved with status: ${result.status}`);
      }
    });
        const sign = extractSignature(data, asyncResult.value.composeType, inDomian,property);
        console.log(
          "getComposeTypeAsync succeeded with composeType: " +
            asyncResult.value.composeType +
            " and coercionType: " +
            asyncResult.value.coercionType
        );
        setSignature(sign);
      } else {
        console.error(asyncResult.error);
      }
      asyncResult.asyncContext.completed();
    });
  }


  export async function setSignatureOnSend(messageType, async) {
    await Office.context.mailbox.item.body.setSignatureAsync(
      messageType,
      {
        coercionType: Office.CoercionType.Html,
      },
      function (asyncResult) {
        async.asyncContext.completed({ allowEvent: true });
      }
    );
  }
  


  export async function setSignature(messageType) {
    await Office.context.mailbox.item.body.setSignatureAsync(
      messageType,
      {
        coercionType: Office.CoercionType.Html,
      },
      function (asyncResult) {
        asyncResult.asyncContext.completed();
      }
    );
  }