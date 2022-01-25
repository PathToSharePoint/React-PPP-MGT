import { GroupType, Login, PeoplePicker, PersonType, TeamsChannelPicker } from '@microsoft/mgt-react';
import * as React from 'react';

import { ICustomPropertyPaneProps } from './ICustomPropertyPaneProps';

import { Provider, teamsTheme, teamsDarkTheme, teamsHighContrastTheme, FormDatepicker, FormRadioGroup } from '@fluentui/react-northstar';
import { PropertyPanePortal } from 'property-pane-portal';
import { Providers } from '@microsoft/mgt';

export const CustomPropertyPane: React.FunctionComponent<ICustomPropertyPaneProps> = (props) => {

  // Teams themes
  let currentTheme;

  switch (props.properties["northstarRadioGroup"]) {
    case "Light": currentTheme = teamsTheme; break;
    case "Dark": currentTheme = teamsDarkTheme; break;
    case "Contrast": currentTheme = teamsHighContrastTheme; break;
    default: currentTheme = teamsTheme;
  }

  return (
    <>
      <Provider theme={currentTheme}>
        <PropertyPanePortal context={props.context}>
          <Login
            data-property="mgtPerson"
            loginCompleted={(e) => {
              console.log("login completed");
              Providers.globalProvider.graph.client.api('me').get()
                .then(getMe => console.log(getMe));
            }}
            logoutCompleted={(e) => { console.log("logout completed"); }}
          />
          {/* <PeoplePicker
            selectionMode="single"
            selectionChanged={(e: any) => {
              console.log(e.detail);
              let users = [];
              e.detail.forEach(dtl => users.push(dtl.userPrincipalName));
              props.updateWebPartProperty("mgtPeoplePicker", users[0]);
            }}
          />
          <PeoplePicker
            data-property="mgtPeoplePicker"
            selectionMode="single"
            defaultSelectedUserIds={[props.properties.mgtPeoplePicker]}
            selectionChanged={(e: any) => {
              console.log(e.detail);
              let users = [];
              e.detail.forEach(dtl => users.push(dtl.userPrincipalName));
              props.updateWebPartProperty("mgtPeoplePicker", users[0]);
            }}
          /> */}
          <PeoplePicker
            data-property="mgtGroupPicker"
            selectionMode="single"
            type={PersonType.group}
            groupType={GroupType.unified}
            defaultSelectedGroupIds={[props.properties.mgtGroupPicker]}
            selectionChanged={(e: any) => {
              console.log(e.detail);
              let users = [];
              e.detail.forEach(dtl => users.push(dtl.userPrincipalName));
              props.updateWebPartProperty("mgtGroupPicker", users[0]);
            }}
          />
          {/* <TeamsChannelPicker
                    data-property="mgtTeamsChannelPicker"
                    defaultValue={[props.properties.mgtTeamsChannelPicker]}
                    selectionChanged={(e: any) => {
                        let slctns = [];
                        console.log(e);
                        e.detail.forEach(dtl => slctns.push(dtl.channel));
                        props.updateWebPartProperty("mgtTeamsChannelPicker", slctns[0]);
                    }}
                /> */}
        </PropertyPanePortal>
      </Provider>
    </>
  );
};