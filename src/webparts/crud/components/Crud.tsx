import * as React from 'react';
//import styles from './Crud2.module.scss';
import { ICrudProps } from './ICrudProps';
import { DatePicker, IDatePickerStrings } from 'office-ui-fabric-react/lib/DatePicker';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { TextField } from 'office-ui-fabric-react/lib/TextField';
import { Label } from 'office-ui-fabric-react/lib/Label';
import { Web} from "@pnp/sp/presets/all";
import { PrimaryButton } from 'office-ui-fabric-react/lib/Button';
import { spfi } from "@pnp/sp";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";


export interface IStates {
    Items: any;
    ID: any;
    EmployeeName: any;
    EmployeeNameId: any;
    HireDate: any;
    JobDescription: any;
    HTML: any;
  }

export default class CRUDReact extends React.Component<ICrudProps, IStates> {
  items: any;
    constructor(props: ICrudProps | Readonly<ICrudProps>) {
        super(props);
        this.state = {
          Items: [],
          EmployeeName: "",
          EmployeeNameId: 0,
          ID: 0,
          HireDate: null,
          JobDescription: "",
          HTML: []
    
        };
      }

      public async componentDidMount() {
        await this.fetchData();
      }

      public async fetchData() {
        //let web = Web(this.props.webURL);
        // const items: any[] = await sp.web.lists.getByTitle("crud").items.get();
        const sp = spfi(this.props.webURL);
        // get all the items from a list
        const items: any[] = await sp.web.lists.getByTitle("crud").items();
        console.log(items);


        this.setState({ Items: items });
        let html = await this.getHTML(items);
        this.setState({ HTML: html });
      }
      public findData = (id: any): void => {
        //this.fetchData();
        var itemID = id;
        var allitems = this.state.Items;
        var allitemsLength = allitems.length;
        if (allitemsLength > 0) {
          for (var i = 0; i < allitemsLength; i++) {
            if (itemID == allitems[i].Id) {
              this.setState({
                ID:itemID,
                EmployeeName:allitems[i].Employee_x0020_Name.Title,
                EmployeeNameId:allitems[i].Employee_x0020_NameId,
                HireDate:new Date(allitems[i].HireDate),
                JobDescription:allitems[i].Job_x0020_Description
              });
          }
        }
      }
      }
      // a l'intérieur de la boucle, 
      // la fonction vérifie si la valeur de "itemID" est égale à la propriété "Id" de l'élément courant dans le tableau.
      // Si les valeurs correspondent, 
      // la fonction utilise la méthode "setState" pour mettre à jour l'état du composant avec
      //  les valeurs suivantes : "ID", "EmployeeName", "EmployeeNameId", "HireDate" et "JobDescription".

      public async getHTML(items: any[]) {
        var tabledata = <table className={ "" /*styles.table*/}> 
          <thead>
            <tr>
              <th>Employee Name</th>
              <th>Hire Date</th>
              <th>Job Description</th>
            </tr>
          </thead>
          <tbody>
            {items && items.map((item, i) => {
              return [
                <tr key={i} onClick={()=>this.findData(item.ID)}>
                  <td>{item.Employee_x0020_Name.Title}</td>
                  <td>{FormatDate(item.HireDate)}</td>
                  <td>{item.Job_x0020_Description}</td>
                </tr>
              ];
            })}
          </tbody>
        </table>;
        return await tabledata;

      }    /* fct renvoie le tableau HTML généré elle prend un tableau d'éléments" en entrée. Le tableau comporte trois colonnes. 
Les lignes du tableau sont générées par mappage sur le tableau "items" et pour chaque élément du tableau, 
une ligne du tableau est créée avec les valeurs de "Employee Name",....
La ligne du tableau a un gestionnaire d'événements de clic "findData" qui est déclenché lorsque la ligne est cliquée,
en passant la propriété "ID" de l'élément en tant que paramètre. 
La fonction "FormatDate" permet de formater la valeur "Hire Date". */

      public _getPeoplePickerItems = async (items: any[]) => {
    
        if (items.length > 0) {
    
          this.setState({ EmployeeName: items[0].text });
          this.setState({ EmployeeNameId: items[0].id });
        }
        else {
          //ID=0;
          this.setState({ EmployeeNameId: "" });
          this.setState({ EmployeeName: "" });
        }
      }
      // public onchange(value, stateValue) {
      //   let state = {};
      //   state[stateValue] = value;
      //   this.setState(state);
      // }
      private async SaveData() {
        let web = Web(this.props.webURL);
        await web.lists.getByTitle("crud").items.add({
    
          Employee_x0020_NameId:this.state.EmployeeNameId,
          HireDate: new Date(this.state.HireDate),
          Job_x0020_Description: this.state.JobDescription,
    
        }).then(i => {
          console.log(i);
        });
        alert("Created Successfully");
        this.setState({EmployeeName:"",HireDate:null,JobDescription:""});
        this.fetchData();
      }
      private async UpdateData() {
        let web = Web(this.props.webURL);
        await web.lists.getByTitle("crud").items.getById(this.state.ID).update({
    
          Employee_x0020_NameId:this.state.EmployeeNameId,
          HireDate: new Date(this.state.HireDate),
          Job_x0020_Description: this.state.JobDescription,
    
        }).then(i => {
          console.log(i);
        });
        alert("Updated Successfully");
        this.setState({EmployeeName:"",HireDate:null,JobDescription:""});
        this.fetchData();
      }
      private async DeleteData() {
        let web = Web(this.props.webURL);
        await web.lists.getByTitle("crud").items.getById(this.state.ID).delete()
        .then(i => {
          console.log(i);
        });
        alert("Deleted Successfully");
        this.setState({EmployeeName:"",HireDate:null,JobDescription:""});
        this.fetchData();
      }
    
      public render(): React.ReactElement<ICrudProps> {
        return (
          <div>
            <h1>CRUD Operations With ReactJs</h1>
            {this.state.HTML}
            <div className={ "" /*styles.btngroup*/}>
              <div><PrimaryButton text="Create" onClick={() => this.SaveData()}/></div>
              <div><PrimaryButton text="Update" onClick={() => this.UpdateData()} /></div>
              <div><PrimaryButton text="Delete" onClick={() => this.DeleteData()}/></div>
            </div>
            <div>
              <form>
                <div>
                  <Label>Employee Name</Label>
                  <PeoplePicker
                    context={this.props.context}
                    personSelectionLimit={1}
                    // defaultSelectedUsers={this.state.EmployeeName===""?[]:this.state.EmployeeName}
                    required={false}
                    onChange={this._getPeoplePickerItems}
                    defaultSelectedUsers={[this.state.EmployeeName?this.state.EmployeeName:""]}
                    showHiddenInUI={false}
                    principalTypes={[PrincipalType.User]}
                    resolveDelay={1000}
                    ensureUser={true}
                  />
                </div>
                <div>
                  <Label>Hire Date</Label>
                  <DatePicker maxDate={new Date()} allowTextInput={false} strings={DatePickerStrings} value={this.state.HireDate} onSelectDate={(e) => { this.setState({ HireDate: e }); }} ariaLabel="Select a date" formatDate={FormatDate} />
                </div>
                <div>
                  <Label>Job Description</Label>
                  { <TextField value={this.state.JobDescription} /* multiline onChange={(value) => this.onchange(value, "JobDescription")} *//> }
                </div>
    
              </form>
            </div>
          </div>
        );
      }
}
export const DatePickerStrings: IDatePickerStrings = {
    months: ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'],
    shortMonths: ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'],
    days: ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'],
    shortDays: ['S', 'M', 'T', 'W', 'T', 'F', 'S'],
    goToToday: 'Go to today',
    prevMonthAriaLabel:'Go to previous month',
    nextMonthAriaLabel:'Go to next month',
    prevYearAriaLabel: 'Go to previous year',
    nextYearAriaLabel: 'Go to next year',
    invalidInputErrorMessage: 'Invalid date format.'
  };
  export const FormatDate = (date:any): string => {
    console.log(date);
    var date1 = new Date(date);
    var year = date1.getFullYear();
    var month = (1 + date1.getMonth()).toString();
    month = month.length > 1 ? month : '0' + month;
    var day = date1.getDate().toString();
    day = day.length > 1 ? day : '0' + day;
    return month + '/' + day + '/' + year;
  };