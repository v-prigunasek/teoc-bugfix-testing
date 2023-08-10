import { ChevronStartIcon } from '@fluentui/react-northstar';
import { ApplicationInsights } from '@microsoft/applicationinsights-web';
import { Client } from "@microsoft/microsoft-graph-client";
import React from 'react';
import Col from 'react-bootstrap/esm/Col';
import Row from 'react-bootstrap/esm/Row';
import "../scss/AdminSettings.module.scss";
import RoleSettings from './RoleSettings';
import { TeamNameConfig } from './TeamNameConfig';
import * as constants from '../common/Constants';

export interface IAdminSettingsProps {
    localeStrings: any;
    onBackClick(showMessageBar: string): void;
    siteId: string;
    graph: Client;
    appInsights: ApplicationInsights;
    userPrincipalName: any;
    showMessageBar(message: string, type: string): void;
    hideMessageBar(): void;
    currentUserDisplayName: string;
    currentUserEmail: string;
    isRolesEnabled: boolean;
    isUserAdmin: boolean;
    configRoleData: any;
    setState: any;
    tenantName: string;
    siteName: any;
    currentThemeName: string;
}

export interface IAdminSettingsState {
    teamNameConfigSettings: boolean;
    roleSettings: boolean;
}

export default class AdminSettings extends React.Component<IAdminSettingsProps, IAdminSettingsState> {
    constructor(props: IAdminSettingsProps) {
        super(props);

        //States
        this.state = {
            teamNameConfigSettings: true,
            roleSettings: false
        }

    }

    //render method
    render() {
        const isDarkOrContrastTheme = this.props.currentThemeName === constants.darkMode || this.props.currentThemeName === constants.contrastMode;
        return (
            <div className='admin-settings'>
                <div className=".col-xs-12 .col-sm-8 .col-md-4 container admin-settings-path">
                    <label>
                        <span
                            onClick={() => this.props.onBackClick("")}
                            onKeyDown={(event) => {
                                if (event.key === constants.enterKey)
                                    this.props.onBackClick("")
                            }} className="go-back">
                            <ChevronStartIcon className="path-back-icon" />
                            <span className="back-label" role="button" tabIndex={0} title="Back">{this.props.localeStrings.back}</span>
                        </span> &nbsp;&nbsp;
                        <span className="right-border">|</span>
                        <span>&nbsp;&nbsp;{this.props.localeStrings.adminSettingsLabel}</span>
                    </label>
                </div>
                <div className={`admin-settings-wrapper${isDarkOrContrastTheme ? " admin-settings-wrapper-darkcontrast" : ""}`}>                            
                    <div className="container">
                        <h1 style={{ "margin": "0" }} aria-live="polite" role="alert"><div className="admin-settings-heading">{this.props.localeStrings.adminSettingsLabel}</div></h1>
                        <Row xl={1} lg={1} md={1} sm={1} xs={1}>
                            <Col md={12}>
                                <div className="toggle-setting-type">
                                    <div
                                        className={`setting-type${this.state.teamNameConfigSettings ? " selected-setting" : ""}`}
                                        onClick={() => this.setState({ teamNameConfigSettings: true, roleSettings: false })}
                                        title={this.props.localeStrings.formTitleTeamNameConfig}
                                    >
                                        {this.props.localeStrings.formTitleTeamNameConfig}
                                    </div>
                                    <div
                                        className={`setting-type${this.state.roleSettings ? " selected-setting" : ""}`}
                                        onClick={() => this.setState({ teamNameConfigSettings: false, roleSettings: true })}
                                        title={this.props.localeStrings.roleSettingsLabel}
                                    >
                                        {this.props.localeStrings.roleSettingsLabel}
                                    </div>
                                </div>
                            </Col>
                        </Row>
                        {this.state.teamNameConfigSettings &&
                            <TeamNameConfig
                                localeStrings={this.props.localeStrings}
                                onBackClick={this.props.onBackClick}
                                siteId={this.props.siteId}
                                graph={this.props.graph}
                                appInsights={this.props.appInsights}
                                userPrincipalName={this.props.userPrincipalName}
                                showMessageBar={this.props.showMessageBar}
                                hideMessageBar={this.props.hideMessageBar}
                                currentThemeName={this.props.currentThemeName}
                            />
                        }
                        {this.state.roleSettings &&
                            <RoleSettings
                                localeStrings={this.props.localeStrings}
                                onBackClick={this.props.onBackClick}
                                currentUserDisplayName={this.props.currentUserDisplayName}
                                currentUserEmail={this.props.currentUserEmail}
                                isRolesEnabled={this.props.isRolesEnabled}
                                isUserAdmin={this.props.isUserAdmin}
                                siteId={this.props.siteId}
                                graph={this.props.graph}
                                configRoleData={this.props.configRoleData}
                                setState={this.props.setState}
                                tenantName={this.props.tenantName}
                                siteName={this.props.siteName}
                                appInsights={this.props.appInsights}
                                userPrincipalName={this.props.userPrincipalName}
                            />
                        }
                    </div>
                </div>
            </div>
        );
    }
}
