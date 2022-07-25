 export function getSignature() {
    var myHeaders = new Headers();
    myHeaders.append("Accept", "application/json");
    var userName=Office.context.mailbox.userProfile.emailAddress;
    return fetch("https://manage.dynasend.net/s/outlook-bundle?email="+userName+"", { method: 'get',headers: myHeaders,redirect: 'follow'})
        .then(response=>response.json());

  }
