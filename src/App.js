import React, { Component } from 'react';
import './App.css';
import { DefaultButton, IButtonProps } from 'office-ui-fabric-react/lib/Button';
import { Link } from 'office-ui-fabric-react/lib/Link';
import { MessageBar, MessageBarType } from 'office-ui-fabric-react/lib/MessageBar';

class App extends Component {
  constructor(props) {
    super(props);

    this.onSetColor = this.onSetColor.bind(this);
  }

  onSetColor() {
    window.Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.format.fill.color = 'green';
      await context.sync();
    });
  }

  render() {
    return (
      <div>
        <MessageBar
          messageBarType={ MessageBarType.success }
          isMultiline={ false }
        >
          Use it like you mean it.  Or some kinda message here using MessageBar Component. <Link href='https://github.com/cpsloal/prototype_office-ui-fabric-react'>See how this was made</Link>
        </MessageBar>
        <div className="Container">
          <img src="assets/logo-filled.png" className="PrototypeLogo" /><h1>Prototype Office UI Fabric with React</h1>
          <p>Click the button below to set the color of the selected cell range to <span className="GreenText">Green.</span></p>
          <DefaultButton
              data-automation-id='test'
              text="Make'em green"
              onClick={this.onSetColor}
              primary={true}
          />
        </div>
      </div>
    );
  }
}

export default App;
