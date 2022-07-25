import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import "./App.css";
import { getSignature } from "../service/APIService/GetSignature";


export default class App extends React.Component {
  constructor() {
    super();
    this.state = {
        isActive:false
    }
}
  componentDidMount() {
    getSignature().then(async (data) => {
  this.setState({isActive: data.use_abbreviated_REPLY_signature});

  localStorage.setItem("AbbreViatedSignature", data.use_abbreviated_REPLY_signature.toString());
});
}


  render() {

    const clicktogetUserName = async () => {
      window.open("https://manage.dynasend.net/mysignature"); 
    };

    const handleChange = (e) => {
      console.log(e)
      var _settings = Office.context.roamingSettings;
      _settings.set("SetFlag",e.target.checked);
      var property = _settings.get("SetFlag");
      Office.context.roamingSettings.saveAsync(function(result) {
        if (result.status !== Office.AsyncResultStatus.Succeeded) {
          console.error(`Action failed with message ${result.error.message}`);
        } else {
          console.log(`Settings saved with status: ${result.status}`);
        }
      });
      console.log(property);
      localStorage.setItem("AbbreViatedSignature", e.target.checked.toString());
      var item_value = localStorage.getItem("AbbreViatedSignature");
      console.log(item_value);
    };

    return (
      <div className="signatureConatiner">
          <>
            <div className="ButtonConatiner">
              <DefaultButton
                className="ms-welcome__action ButtonClass"
                iconProps={{ iconName: "ChevronRight" }}
                onClick={clicktogetUserName}
              >
                Create / Edit Signature
              </DefaultButton>
            </div>

            <div className="validationContainer">
              <input
                type="checkbox"
                id="AbbSign"
                onChange={(e) => handleChange(e)}
                disabled={this.state.isActive}
              />
              use abbreviated REPLY signature
            </div>
          </>
      </div>
    );
  }
}

export { App };
