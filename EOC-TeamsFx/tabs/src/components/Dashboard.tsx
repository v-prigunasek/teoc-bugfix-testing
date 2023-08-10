import { Callout, DirectionalHint, Icon, IPivotItemProps, Pivot, PivotItem } from '@fluentui/react';
import { Button, Flex, FormInput, Loader, SearchIcon } from "@fluentui/react-northstar";
import { ApplicationInsights } from '@microsoft/applicationinsights-web';
import { Client } from "@microsoft/microsoft-graph-client";
import * as microsoftTeams from "@microsoft/teams-js";
import 'bootstrap/dist/css/bootstrap.min.css';
import * as React from "react";
import BootstrapTable from "react-bootstrap-table-next";
import 'react-bootstrap-table-next/dist/react-bootstrap-table2.min.css';
import paginationFactory from 'react-bootstrap-table2-paginator';
import CommonService from "../common/CommonService";
import * as constants from '../common/Constants';
import * as graphConfig from '../common/graphConfig';
import siteConfig from '../config/siteConfig.json';
import '../scss/Dashboard.module.scss';

export interface IDashboardProps {
    graph: Client;
    tenantName: string;
    siteId: string;
    onCreateTeamClick: Function;
    onEditButtonClick(incidentData: any): void;
    localeStrings: any;
    onBackClick(showMessageBar: string): void;
    showMessageBar(message: string, type: string): void;
    hideMessageBar(): void;
    appInsights: ApplicationInsights;
    userPrincipalName: any;
    siteName: any;
    onShowAdminSettings: Function;
    onShowIncidentHistory: Function;
    onShowActiveBridge: Function;
    isRolesEnabled: boolean;
    isUserAdmin: boolean;
    settingsLoader: boolean;
}

export interface IDashboardState {
    allIncidents: any;
    planningIncidents: any;
    activeIncidents: any;
    completedIncidents: any;
    filteredAllIncidents: any;
    filteredPlanningIncidents: any;
    filteredActiveIncidents: any;
    filteredCompletedIncidents: any;
    searchText: string | undefined;
    isDesktop: boolean;
    showLoader: boolean;
    loaderMessage: string;
    selectedIncident: any;
    isManageCalloutVisible: boolean;
    currentTab: any;
    incidentIdAriaSort: any;
    incidentNameAriaSort: any;
    locationAriaSort: any;
}

// interface for Dashboard fields
export interface IListItem {
    itemId: string;
    incidentId: string;
    incidentName: string;
    incidentCommander: string;
    status: string;
    location: string;
    startDate: string;
}

class Dashboard extends React.PureComponent<IDashboardProps, IDashboardState> {
    private dashboardRef: React.RefObject<HTMLDivElement>;
    constructor(props: IDashboardProps) {
        super(props);
        this.dashboardRef = React.createRef();
        this.state = {
            allIncidents: [],
            planningIncidents: [],
            activeIncidents: [],
            completedIncidents: [],
            filteredAllIncidents: [],
            filteredPlanningIncidents: [],
            filteredActiveIncidents: [],
            filteredCompletedIncidents: [],
            searchText: "",
            isDesktop: true,
            showLoader: true,
            loaderMessage: this.props.localeStrings.genericLoaderMessage,
            selectedIncident: [],
            isManageCalloutVisible: false,
            currentTab: "",
            incidentIdAriaSort: "",
            incidentNameAriaSort: "",
            locationAriaSort: "",
        };

        this.actionFormatter = this.actionFormatter.bind(this);
        this.renderIncidentSettings = this.renderIncidentSettings.bind(this);
    }

    private dataService = new CommonService();

    // set the state object for screen size
    resize = () => this.setState({ isDesktop: window.innerWidth > constants.mobileWidth });

    // before unmounting, remove event listener
    componentWillUnmount() {
        window.removeEventListener("resize", this.resize.bind(this));
    }

    // component life cycle method
    public async componentDidMount() {
        window.addEventListener("resize", this.resize.bind(this));
        this.resize();
        // Get dashboard data
        this.getDashboardData();
    }


    // connect with servie layer to get all incidents
    getDashboardData = async () => {

        this.setState({
            showLoader: true
        })

        try {
            // create graph endpoint for querying Incident Transaction list
            let graphEndpoint = `${graphConfig.spSiteGraphEndpoint}${this.props.siteId}/lists/${siteConfig.incidentsList}/items?$expand=fields
                ($select=StatusLookupId,Status,id,IncidentId,IncidentName,IncidentCommander,Location,StartDateTime,
                Modified,TeamWebURL,Description,IncidentType,RoleAssignment,RoleLeads,Severity,PlanID,
                BridgeID,BridgeLink,NewsTabLink,CloudStorageLink)&$Top=5000`;

            const allIncidents = this.sortDashboardData(await this.dataService.getDashboardData(graphEndpoint, this.props.graph));
            console.log(constants.infoLogPrefix + "All Incidents retrieved");

            // filter for Planning tab
            const planningIncidents = allIncidents.filter((e: any) => e.incidentStatusObj.status === constants.planning);

            // filter for Active tab
            const activeIncidents = allIncidents.filter((e: any) => e.incidentStatusObj.status === constants.active);

            // filter for Completed tab
            const completedIncidents = allIncidents.filter((e: any) => e.incidentStatusObj.status === constants.closed);

            this.setState({
                allIncidents: allIncidents,
                planningIncidents: planningIncidents,
                activeIncidents: activeIncidents,
                completedIncidents: completedIncidents,
                filteredAllIncidents: [...allIncidents],
                filteredPlanningIncidents: [...planningIncidents],
                filteredCompletedIncidents: [...completedIncidents],
                filteredActiveIncidents: [...activeIncidents],
                showLoader: false
            })
        } catch (error: any) {
            this.setState({
                showLoader: false
            })
            console.error(
                constants.errorLogPrefix + "_Dashboard_GetDashboardData \n",
                JSON.stringify(error)
            );
            this.props.showMessageBar(this.props.localeStrings.genericErrorMessage + ", " + this.props.localeStrings.getIncidentsFailedErrMsg, constants.messageBarType.error);
            // Log Exception
            this.dataService.trackException(this.props.appInsights, error, constants.componentNames.DashboardComponent, 'GetDashboardData', this.props.userPrincipalName);
        }
    }

    // sort data based on order
    sortDashboardData = (allIncidents: any): any => {
        return allIncidents.sort((currIncident: any, prevIncident: any) => (parseInt(currIncident.itemId) < parseInt(prevIncident.itemId)) ? 1 : ((parseInt(prevIncident.itemId) < parseInt(currIncident.itemId)) ? -1 : 0));
    }

    // update icon on pivots(tabs)
    _customRenderer(
        link?: IPivotItemProps,
        defaultRenderer?: (link?: IPivotItemProps) => JSX.Element | null,
    ): JSX.Element | null {
        if (!link || !defaultRenderer) {
            return null;
        }
        return (
            <span>
                <img src={require(`../assets/Images/${link.itemKey}ItemsSelected.svg`)} alt={`${link.headerText}`} id="pivot-icon-selected" />
                <img src={require(`../assets/Images/${link.itemKey}Items.svg`)} alt={`${link.headerText}`} id="pivot-icon" />
                <span id="state">&nbsp;&nbsp;{link.headerText}&nbsp;&nbsp;</span>
                <span id="count">|&nbsp;{link.itemCount}</span>
                {defaultRenderer({ ...link, headerText: undefined, itemCount: undefined })}
            </span>
        );
    }


    //pagination properties for bootstrap table
    private pagination = paginationFactory({
        page: 1,
        sizePerPage: constants.dashboardPageSize,
        showTotal: true,
        alwaysShowAllBtns: false,
        //customized the render options for pagesize button in the pagination for accessbility
        sizePerPageRenderer: ({
            options,
            currSizePerPage,
            onSizePerPageChange
        }) => (
            <div className="btn-group" role="group">
                {
                    options.map((option) => {
                        const isSelect = currSizePerPage === `${option.page}`;
                        return (
                            <button
                                key={option.text}
                                type="button"
                                onClick={() => onSizePerPageChange(option.page)}
                                className={`btn ${isSelect ? 'sizeperpage-selected' : 'sizeperpage'}`}
                                aria-label={constants.sizePerPageLabel + option.text}
                            >
                                {option.text}
                            </button>
                        );
                    })
                }
            </div>
        ),
        //customized the render options for page list in the pagination for accessbility
        pageButtonRenderer: (options: any) => {
            const handleClick = (e: any) => {
                e.preventDefault();
                if (options.disabled) return;
                options.onPageChange(options.page);
            };
            const className = `${options.active ? 'active ' : ''}${options.disabled ? 'disabled ' : ''}`;
            let ariaLabel = "";
            let pageText = "";
            switch (options.title) {
                case "first page":
                    ariaLabel = `Go to ${options.title}`;
                    pageText = '<<';
                    break;
                case "previous page":
                    ariaLabel = `Go to ${options.title}`;
                    pageText = '<';
                    break;
                case "next page":
                    ariaLabel = `Go to ${options.title}`;
                    pageText = '>';
                    break;
                case "last page":
                    ariaLabel = `Go to ${options.title}`;
                    pageText = '>>';
                    break;
                default:
                    ariaLabel = `Go to page ${options.title}`;
                    pageText = options.title;
                    break;
            }
            return (
                <li key={options.title} className={`${className}page-item`} role="presentation" title={ariaLabel}>
                    <a className="page-link" href="#" onClick={handleClick} role="button" aria-label={ariaLabel}>
                        <span aria-hidden="true">{pageText}</span>
                    </a>
                </li>
            );
        },
        //customized the page total renderer in the pagination for accessbility
        paginationTotalRenderer: (from, to, size) => {
            const resultsFound = size !== 0 ? `Showing ${from} to ${to} of ${size} Results` : ""
            return (
                <span className="react-bootstrap-table-pagination-total" aria-live="polite" role="status">
                    {resultsFound}
                </span>
            )
        }
    });

    // filter incident based on entered keyword
    searchDashboard = (searchText: any) => {
        let searchKeyword = "";
        // convert to lower case
        if (searchText.target.value) {
            searchKeyword = searchText.target.value.toLowerCase();
        }
        const allIncidents = this.state.allIncidents;
        let filteredAllIncidents = allIncidents.filter((incident: any) => {
            return ((incident["incidentName"] && incident["incidentName"].toLowerCase().indexOf(searchKeyword) > -1) ||
                (incident["incidentId"] && (incident["incidentId"]).toString().toLowerCase().indexOf(searchKeyword) > -1) ||
                (incident["incidentCommander"] && incident["incidentCommander"].toLowerCase().indexOf(searchKeyword) > -1) ||
                (incident["location"] && incident["location"].toLowerCase().indexOf(searchKeyword) > -1))
        });

        //On Click of Cancel icon
        if (searchText.cancelable) {
            filteredAllIncidents = this.state.allIncidents;
        }
        this.setState({
            searchText: searchText.target.value,
            filteredAllIncidents: filteredAllIncidents,
            filteredPlanningIncidents: filteredAllIncidents.filter((e: any) => e.status === constants.planning),
            filteredActiveIncidents: filteredAllIncidents.filter((e: any) => e.status === constants.active),
            filteredCompletedIncidents: filteredAllIncidents.filter((e: any) => e.status === constants.closed),
        });
    }

    // format the cell for Incident ID column to fix accessibility issues
    incidentIdFormatter = (cell: any, gridRow: any, rowIndex: any, formatExtraData: any) => {
        const ariaLabel = `Row ${rowIndex + 2} ${this.props.localeStrings.incidentId} ${cell}`
        return (
            <span
                aria-label={ariaLabel}
                tabIndex={0}
                role="link"
                className="incIdDeepLink"
                onClick={() => this.onDeepLinkClick(gridRow)}
                onKeyDown={(event) => {
                    if (event.key === constants.enterKey)
                        this.onDeepLinkClick(gridRow)
                }}
            >{cell}</span>
        );
    }

    // format the cell for Incident Name column to fix accessibility issues
    incidentNameFormatter = (cell: any, gridRow: any, rowIndex: any, formatExtraData: any) => {
        const ariaLabel = `${this.props.localeStrings.incidentName} ${cell}`
        return (
            <span aria-label={ariaLabel}>{cell}</span>
        );
    }

    // format the cell for Severity column to fix accessibility issues
    severityFormatter = (cell: any, gridRow: any, rowIndex: any, formatExtraData: any) => {
        const ariaLabel = `${this.props.localeStrings.fieldSeverity} ${cell}`
        return (
            <span aria-label={ariaLabel}>{cell}</span>
        );
    }

    // format the cell for Incident Commander column to fix accessibility issues
    incidentCommanderFormatter = (cell: any, gridRow: any, rowIndex: any, formatExtraData: any) => {
        const ariaLabel = `${this.props.localeStrings.incidentCommander} ${cell}`
        return (
            <span aria-label={ariaLabel}>{cell}</span>
        );
    }

    // format the cell for Status column to fix accessibility issues
    statusFormatter = (cell: any, row: any, rowIndex: any, formatExtraData: any) => {
        if (row.incidentStatusObj.status === constants.closed) {
            return (
                <img src={require("../assets/Images/ClosedIcon.svg").default} aria-label={`${this.props.localeStrings.status} ${this.props.localeStrings.closed}`} className="status-icon" />
            );
        }
        if (row.incidentStatusObj.status === constants.active) {
            return (
                <img src={require("../assets/Images/ActiveIcon.svg").default} aria-label={`${this.props.localeStrings.status} ${this.props.localeStrings.active}`} className="status-icon" />
            );
        }
        if (row.incidentStatusObj.status === constants.planning) {
            return (
                <img src={require("../assets/Images/PlanningIcon.svg").default} aria-label={`${this.props.localeStrings.status} ${this.props.localeStrings.planning}`} className="status-icon" />
            );
        }
    };

    // format the cell for Location column to fix accessibility issues
    locationFormatter = (cell: any, gridRow: any, rowIndex: any, formatExtraData: any) => {
        const ariaLabel = `${this.props.localeStrings.location} ${cell}`
        return (
            <span aria-label={ariaLabel}>{cell}</span>
        );
    }

    // format the cell for Start Date Time column to fix accessibility issues
    startDateTimeFormatter = (cell: any, gridRow: any, rowIndex: any, formatExtraData: any) => {
        const ariaLabel = `${this.props.localeStrings.startDate} ${cell}`
        return (
            <span aria-label={ariaLabel}>{cell}</span>
        );
    }

    // format the cell for Action column to fix accessibility issues
    public actionFormatter(_cell: any, gridRow: any, _rowIndex: any, _formatExtraData: any) {
        return (
            <span>
                {/* active dashboard icon in dashboard, on click will navigate to edit incident form */}
                <img
                    src={require("../assets/Images/ActiveBridgeIcon.svg").default}
                    aria-label={`${this.props.localeStrings.action} ${this.props.localeStrings.activeDashboard}`}
                    className="grid-active-bridge-icon"
                    onClick={() => this.props.onShowActiveBridge(gridRow)}
                    onKeyDown={(event) => {
                        if (event.key === constants.enterKey)
                            this.props.onShowActiveBridge(gridRow)
                    }}
                    tabIndex={0}
                    role="button"
                />
                {/* edit icon in dashboard, on click will navigate to edit incident form */}
                <img
                    src={require("../assets/Images/GridEditIcon.svg").default}
                    aria-label={`${this.props.localeStrings.action} ${this.props.localeStrings.edit}`}
                    className="grid-edit-icon"
                    onClick={() => this.props.onEditButtonClick(gridRow)}
                    onKeyDown={(event) => {
                        if (event.key === constants.enterKey)
                            this.props.onEditButtonClick(gridRow)
                    }}
                    tabIndex={0}
                    role="button"
                />

                {/* view version history icon in dashboard, on click will navigate to incident history form */}
                <img
                    src={require("../assets/Images/IncidentHistoryIcon.svg").default}
                    aria-label={`${this.props.localeStrings.action} ${this.props.localeStrings.viewIncidentHistory}`}
                    className="grid-version-history-icon"
                    onClick={() => this.props.onShowIncidentHistory(gridRow.incidentId)}
                    onKeyDown={(event) => {
                        if (event.key === constants.enterKey)
                            this.props.onShowIncidentHistory(gridRow.incidentId)
                    }}
                    tabIndex={0}
                    role="button"
                />
            </span>
        );
    }


    // create deep link to open the associated Team
    onDeepLinkClick = (rowData: any) => {
        microsoftTeams.executeDeepLink(rowData.teamWebURL);
    }

    //Incident Settings Area
    public renderIncidentSettings = () => {
        // added constants for manage dropdown menu
        const manageButtonId = 'manage-callout-button';
        const manageLabelId = 'manage-callout-label';
        const manageDescriptionId = 'manage-callout-description';
        return (
            <Flex space="between" wrap={true}>
                <div
                    className={`manage-links${this.state.isManageCalloutVisible ? " callout-visible" : ""}`}
                    onClick={() => this.setState({ isManageCalloutVisible: !this.state.isManageCalloutVisible })}
                    id={manageButtonId}
                    tabIndex={0}
                    onKeyDown={(event) => {
                        if (event.key === constants.enterKey)
                            this.setState({ isManageCalloutVisible: !this.state.isManageCalloutVisible })
                    }}
                    role="button"
                    title="Manage"
                >
                    <img
                        src={require("../assets/Images/ManageIcon.svg").default}
                        className="manage-icon"
                        alt=""
                    />
                    <img
                        src={require("../assets/Images/ManageIconActive.svg").default}
                        className='manage-icon-active'
                        alt=""
                    />
                    <div className='manage-label'>{this.props.localeStrings.manageLabel}</div>
                    {this.state.isManageCalloutVisible ?
                        <Icon iconName="ChevronUp" />
                        :
                        <Icon iconName="ChevronDown" />
                    }
                </div>

                {this.state.isManageCalloutVisible ? (
                    <Callout
                        ariaLabelledBy={manageLabelId}
                        ariaDescribedBy={manageDescriptionId}
                        role="menu"
                        className="manage-links-callout"
                        gapSpace={10}
                        target={`#${manageButtonId}`}
                        isBeakVisible={false}
                        onDismiss={() => this.setState({ isManageCalloutVisible: false })}
                        directionalHint={DirectionalHint.bottomLeftEdge}
                        onPositioned={() => document?.getElementsByClassName('manage-links-callout')[0]
                            ?.getElementsByTagName("div")[0]?.setAttribute("aria-busy", "true")}
                        setInitialFocus
                    >
                        <div>
                            <div title={this.props.localeStrings.manageIncidentTypesTooltip} className="dashboard-link">
                                <a title={this.props.localeStrings.manageIncidentTypesTooltip} href={`https://${this.props.tenantName}/sites/${this.props.siteName}/lists/${siteConfig.incTypeList}`} target='_blank' rel="noreferrer">
                                    <img src={require("../assets/Images/Manage Incident Types.svg").default} alt="" />
                                    <span role="button" className="manage-callout-text">{this.props.localeStrings.incidentTypesLabel}</span>
                                </a>
                            </div>
                            <div title={this.props.localeStrings.manageRolesTooltip} className="dashboard-link">
                                <a title={this.props.localeStrings.manageRolesTooltip} href={`https://${this.props.tenantName}/sites/${this.props.siteName}/lists/${siteConfig.roleAssignmentList}`} target='_blank' rel="noreferrer">
                                    <img src={require("../assets/Images/Manage Roles.svg").default} alt="" />
                                    <span role="button" className="manage-callout-text">{this.props.localeStrings.roles}</span>
                                </a>
                            </div>
                            <div
                                className="dashboard-link"
                                title={this.props.localeStrings.adminSettingsLabel}
                                onClick={() => this.props.onShowAdminSettings()}
                                onKeyDown={(event) => {
                                    if (event.key === constants.enterKey)
                                        this.props.onShowAdminSettings()
                                    else if (!event.shiftKey)
                                        this.setState({ isManageCalloutVisible: false })
                                }}
                                role="button"
                            >
                                <span className="team-name-link" tabIndex={0}>
                                    <img
                                        src={require("../assets/Images/AdminSettings.svg").default}
                                        alt=""
                                    />
                                    <span className="manage-callout-text">
                                        {this.props.localeStrings.adminSettingsLabel}
                                    </span>
                                </span>
                            </div>
                        </div>
                    </Callout>
                ) : null}
                <Button
                    primary id="create-incident-btn"
                    fluid={true}
                    onClick={() => this.props.onCreateTeamClick()}
                    title={this.props.localeStrings.btnCreateIncident}
                >
                    <img src={require("../assets/Images/ButtonEditIcon.svg").default} alt="edit icon" />
                    {this.props.localeStrings.btnCreateIncident}
                </Button>
            </Flex>
        );
    }

    //render the sort caret on the header column for accessbility
    customSortCaret = (order: any, column: any) => {
        const ariaLabel = navigator.userAgent.match(/iPhone/i) ? "sortable" : "";
        const id = column.dataField;
        if (!order) {
            return (
                <span className="sort-order" id={id} aria-label={ariaLabel}>
                    <span className="dropdown-caret">
                    </span>
                    <span className="dropup-caret">
                    </span>
                </span>);
        }
        else if (order === 'asc') {
            column.dataField === "incidentId" ?
                this.setState({ incidentIdAriaSort: constants.sortAscAriaSort, incidentNameAriaSort: "", locationAriaSort: "" }) :
                column.dataField === "incidentName" ?
                    this.setState({ incidentNameAriaSort: constants.sortAscAriaSort, incidentIdAriaSort: "", locationAriaSort: "" }) :
                    this.setState({ locationAriaSort: constants.sortAscAriaSort, incidentNameAriaSort: "", incidentIdAriaSort: "" })
            return (
                <span className="sort-order">
                    <span className="dropup-caret">
                    </span>
                </span>);
        }
        else if (order === 'desc') {
            column.dataField === "incidentId" ?
                this.setState({ incidentIdAriaSort: constants.sortDescAriaSort, incidentNameAriaSort: "", locationAriaSort: "" }) :
                column.dataField === "incidentName" ?
                    this.setState({ incidentNameAriaSort: constants.sortDescAriaSort, incidentIdAriaSort: "", locationAriaSort: "" }) :
                    this.setState({ locationAriaSort: constants.sortDescAriaSort, incidentNameAriaSort: "", incidentIdAriaSort: "" })
            return (
                <span className="sort-order">
                    <span className="dropdown-caret">
                    </span>
                </span>);
        }
        return null;
    }

    //custom header format for sortable column for accessbility
    headerFormatter(column: any, colIndex: any, { sortElement, filterElement }: any) {
        const id = column.dataField;
        //adding sortable information to aria-label to fix the accessibility issue in iOS Voiceover
        if (navigator.userAgent.match(/iPhone/i)) {
            return (
                <button tabIndex={-1} aria-describedby={id} aria-label={column.text} className='sort-header'>
                    {column.text}
                    {sortElement}
                </button>
            );
        }
        else {
            return (
                <>
                    {column.text}
                    {sortElement}
                </>
            );
        }
    }


    public render() {
        // Header object for dashboard
        const dashboardHeader = [
            {
                dataField: 'incidentId',
                text: this.props.localeStrings.incidentId,
                sort: true,
                sortCaret: this.customSortCaret,
                formatter: this.incidentIdFormatter,
                headerFormatter: this.headerFormatter,
                headerAttrs: { 'aria-sort': this.state.incidentIdAriaSort, 'role': 'columnheader', 'scope': 'col' },
                attrs: { 'role': 'presentation' }

            }, {
                dataField: 'incidentName',
                text: this.props.localeStrings.incidentName,
                sort: true,
                sortCaret: this.customSortCaret,
                formatter: this.incidentNameFormatter,
                headerFormatter: this.headerFormatter,
                headerAttrs: { 'aria-sort': this.state.incidentNameAriaSort, 'role': 'columnheader', 'scope': 'col' },
                attrs: { 'role': 'presentation' }

            }, {
                dataField: 'severity',
                text: this.props.localeStrings.fieldSeverity,
                formatter: this.severityFormatter,
                headerAttrs: { 'role': 'columnheader', 'scope': 'col' },
                attrs: { 'role': 'presentation' }

            }, {
                dataField: 'incidentCommander',
                text: this.props.localeStrings.incidentCommander,
                formatter: this.incidentCommanderFormatter,
                headerAttrs: { 'role': 'columnheader', 'scope': 'col' },
                attrs: { 'role': 'presentation' }

            }, {
                dataField: 'status',
                text: this.props.localeStrings.status,
                formatter: this.statusFormatter,
                headerAttrs: { 'role': 'columnheader', 'scope': 'col' },
                attrs: { 'role': 'presentation' }

            }, {
                dataField: 'location',
                text: this.props.localeStrings.location,
                sort: true,
                sortCaret: this.customSortCaret,
                headerFormatter: this.headerFormatter,
                formatter: this.locationFormatter,
                headerAttrs: { 'aria-sort': this.state.locationAriaSort, 'role': 'columnheader', 'scope': 'col' },
                attrs: { 'role': 'presentation' }

            }, {
                dataField: 'startDate',
                text: this.props.localeStrings.startDate,
                formatter: this.startDateTimeFormatter,
                headerAttrs: { 'role': 'columnheader', 'scope': 'col' },
                attrs: { 'role': 'presentation' }

            }, {
                dataField: 'action',
                text: this.props.localeStrings.action,
                formatter: this.actionFormatter,
                headerAttrs: { 'role': 'columnheader', 'scope': 'col' },
                attrs: { 'role': 'presentation' }

            }
        ]

        return (
            <>
                {this.state.showLoader ?
                    <>
                        <Loader label={this.state.loaderMessage} size="largest" />
                    </>
                    :
                    <div>
                        <div id="dashboard-search-btn-area">
                            <div className="container">
                                <Flex space="between" wrap={true}>
                                    <div className="search-area">
                                        <FormInput
                                            type="text"
                                            icon={<SearchIcon />}
                                            clearable
                                            placeholder={this.props.localeStrings.searchPlaceholder}
                                            fluid={true}
                                            maxLength={constants.maxCharLengthForSingleLine}
                                            required
                                            title={this.props.localeStrings.searchPlaceholder}
                                            onChange={(evt) => this.searchDashboard(evt)}
                                            value={this.state.searchText}
                                            successIndicator={false}
                                            aria-describedby='noincident-all-tab noincident-active-tab noincident-planning-tab noincident-completed-tab'
                                        />
                                    </div>
                                    {this.props.isRolesEnabled ?
                                        this.props.isUserAdmin ? this.renderIncidentSettings() : <></>
                                        : this.props.settingsLoader ? <Loader size="smallest" className="settings-loader" /> : this.renderIncidentSettings()
                                    }
                                </Flex>
                            </div>
                        </div>
                        <div id="dashboard-pivot-container">
                            <div className="container">
                                <h1><div className="grid-heading">{this.props.localeStrings.incidentDetails}</div></h1>
                                <Flex gap={this.state.isDesktop ? "gap.medium" : "gap.small"} space="evenly" id="status-indicators" wrap={true}>
                                    <Flex gap={this.state.isDesktop ? "gap.small" : "gap.smaller"}>
                                        <img src={require("../assets/Images/AllItems.svg").default} alt="All Items" />
                                        <label>{this.props.localeStrings.all}</label>
                                    </Flex>
                                    <Flex gap={this.state.isDesktop ? "gap.small" : "gap.smaller"}>
                                        <img src={require("../assets/Images/PlanningItems.svg").default} alt="Planning Items" />
                                        <label>{this.props.localeStrings.planning}</label>
                                    </Flex>
                                    <Flex gap={this.state.isDesktop ? "gap.small" : "gap.smaller"}>
                                        <img src={require("../assets/Images/ActiveItems.svg").default} alt="Active Items" />
                                        <label>{this.props.localeStrings.active}</label>
                                    </Flex>
                                    <Flex gap={this.state.isDesktop ? "gap.small" : "gap.smaller"}>
                                        <img src={require("../assets/Images/ClosedItems.svg").default} alt="Completed Items" />
                                        <label>{this.props.localeStrings.closed}</label>
                                    </Flex>
                                </Flex>
                                <Pivot
                                    aria-label="Incidents Details"
                                    linkFormat="tabs"
                                    overflowBehavior='none'
                                    id="piv-tabs"
                                    onLinkClick={(item, ev) => (this.setState({ currentTab: item?.props.headerText }))}
                                    ref={this.dashboardRef}
                                >
                                    <PivotItem
                                        headerText={this.props.localeStrings.all}
                                        itemCount={this.state.filteredAllIncidents.length}
                                        itemKey="All"
                                        onRenderItemLink={this._customRenderer}
                                    >
                                        <BootstrapTable
                                            striped
                                            bootstrap4
                                            bordered={false}
                                            keyField="incidentId"
                                            columns={dashboardHeader}
                                            data={this.state.filteredAllIncidents}
                                            pagination={this.pagination}
                                            noDataIndication={() => (<div id="noincident-all-tab" aria-live="polite" role="status">{this.props.localeStrings.noIncidentsFound}</div>)}
                                        />
                                    </PivotItem>
                                    <PivotItem
                                        headerText={this.props.localeStrings.planning}
                                        itemCount={this.state.filteredPlanningIncidents.length}
                                        itemKey="Planning"
                                        onRenderItemLink={this._customRenderer}
                                    >
                                        <BootstrapTable
                                            striped
                                            bootstrap4
                                            bordered={false}
                                            keyField="incidentId"
                                            columns={dashboardHeader}
                                            data={this.state.filteredPlanningIncidents}
                                            pagination={this.pagination}
                                            noDataIndication={() => (<div id="noincident-planning-tab" aria-live="polite" role="status">{this.props.localeStrings.noIncidentsFound}</div>)}
                                        />
                                    </PivotItem>
                                    <PivotItem
                                        headerText={this.props.localeStrings.active}
                                        itemCount={this.state.filteredActiveIncidents.length}
                                        itemKey="Active"
                                        onRenderItemLink={this._customRenderer}
                                    >
                                        <BootstrapTable
                                            striped
                                            bootstrap4
                                            bordered={false}
                                            keyField="incidentId"
                                            columns={dashboardHeader}
                                            data={this.state.filteredActiveIncidents}
                                            pagination={this.pagination}
                                            noDataIndication={() => (<div id="noincident-active-tab" aria-live="polite" role="status">{this.props.localeStrings.noIncidentsFound}</div>)}
                                        />
                                    </PivotItem>
                                    <PivotItem
                                        headerText={this.props.localeStrings.closed}
                                        itemCount={this.state.filteredCompletedIncidents.length}
                                        itemKey="Closed"
                                        onRenderItemLink={this._customRenderer}
                                    >
                                        <BootstrapTable
                                            striped
                                            bootstrap4
                                            bordered={false}
                                            keyField="incidentId"
                                            columns={dashboardHeader}
                                            data={this.state.filteredCompletedIncidents}
                                            pagination={this.pagination}
                                            noDataIndication={() => (<div id="noincident-completed-tab" aria-live="polite" role="status">{this.props.localeStrings.noIncidentsFound}</div>)}
                                        />
                                    </PivotItem>
                                </Pivot>
                            </div>
                        </div>
                    </div>
                }
            </>
        );
    }
}

export default Dashboard;
