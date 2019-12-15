import { sp } from "@pnp/sp";

export class EmailProperties {
    public To: string[];
    public CC?: string[];
    public BCC?: string[];
    public Subject: string;
    public Body: string;
    public From?: string;
}
export default class DataService {

    private _employeeDetails: any;
    public get employeeDetails(): any {
        return this._employeeDetails;
    }
    public set employeeDetails(v: any) {
        this._employeeDetails = v;
    }


    private _emailTemplateItems: any;
    public get emailTemplateItems(): any {
        return this._emailTemplateItems;
    }
    public set emailTemplateItems(v: any) {
        this._emailTemplateItems = v;
    }

    constructor() {
        this.employeeDetails = [];
        this.fetchAllEmailTemplateItems();
    }
    public randomNumber(min: number, max: number): any {
        return Math.floor(Math.random() * (max - min) + min);
    }
    public generateSecrateSantaAssignedArray(employeeDetails): any {
        employeeDetails.map((value, index) => {
            let temparray = employeeDetails.filter((val) => { return val.isSanta === false; });
            let random = index;
            do {
                random = this.randomNumber(1, temparray.length);
            } while (random === index);
            let indice = -1;
            if (employeeDetails.length - 1 == index) {
                random--;
            }
            employeeDetails.some((el, tempindex) => {
                if (el.name == temparray[random].name) {
                    indice = tempindex;
                    return true;
                }
            });
            employeeDetails[indice].isSanta = true;
            employeeDetails[index].asssigned = temparray[random].name;
        });
        this.employeeDetails = employeeDetails;
        return this.employeeDetails;
    }
    public setIsSantaFalse(employeeDetails): any {
        employeeDetails.map((val, index) => {
            employeeDetails[index].isSanta = false;
        });
    }
    public async getEmployeeList(listID: string): Promise<any> {
        return new Promise<any>((resolve, reject) => {
            sp.web.lists.getById(listID).items.select('Title,Employee/Title,Employee/EMail,ID').expand('Employee').get().then((response) => {
                console.log(response);
                let tempArry: any = [];
                response.map((val, index) => {
                    let tempObject = { ID: val.ID, name: val.Employee.Title, isSanta: false, asssigned: "", email: val.Employee.EMail };
                    tempArry.push(tempObject);
                });
                this.employeeDetails = tempArry;
                resolve(true);
            }).catch((error) => {
                console.log(error);
                reject(true);
            });
        });
    }

    public fetchAllEmailTemplateItems() {
        sp.web.lists.getByTitle('Email Template').items.get().then(response => {
            console.log(response);
            this.emailTemplateItems = response;
        }).catch(error => console.log(error));
    }
    public sendEmail(): any {
        let temparray = this.emailTemplateItems.filter((val) => { return val.Title === "SantaAllocation";});
        if(temparray.length>0){
            this.employeeDetails.map((val, index) => {
                let emailProperties: EmailProperties = new EmailProperties;
                emailProperties.To = [val.email];
                let body = temparray[0].Body.replace("{{santaName}}",val.asssigned);
                emailProperties.Body = body;
                emailProperties.Subject = temparray[0].Subject;
                sp.utility.sendEmail(emailProperties);
            });
        }else{
            this.employeeDetails.map((val, index) => {
                let emailProperties: EmailProperties = new EmailProperties;
                emailProperties.To = [val.email];
                emailProperties.Body = `You have been Assigned ${val.asssigned} User as Secret Santa`;
                emailProperties.Subject = `Secret Santa is Assigned`;
                sp.utility.sendEmail(emailProperties);
            });
        }        
    }
    public async fetchExistingWishValue(currentUserEmail: string): Promise<any> {
        return new Promise<any>((resolve, reject) => {
            sp.web.ensureUser(currentUserEmail).then(result => {
                sp.web.lists.getByTitle('wish list').items.select('Author/ID,ID,Wish').expand('Author').filter(`Author/ID eq ${result.data.Id}`).get().then((items) => {
                    resolve(items);
                }).catch((error) => {
                    reject(false);
                });
            });
        });
    }

    public async addupdateWishValue(wishValue: string, itemId?: number): Promise<any> {
        return new Promise<any>((resolve, reject) => {
            if (!itemId) {
                sp.web.lists.getByTitle('wish list').items.add({
                    Wish: wishValue
                }).then((result) => {
                    resolve(result);
                }).catch((error) => {
                    reject(false);
                });
            } else {
                sp.web.lists.getByTitle('wish list').items.getById(itemId).update({
                    Wish: wishValue
                }).then((items) => {
                    resolve(items);
                }).catch((error) => {
                    reject(false);
                });
            }
        });
    }

}