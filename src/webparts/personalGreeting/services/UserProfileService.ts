import { WebPartContext } from "@microsoft/sp-webpart-base";
import { getSP } from "../pnpjsConfig"
import { SPFI } from "@pnp/sp";
import { Logger, LogLevel } from "@pnp/logging";
import "@pnp/sp/webs";
import "@pnp/sp/site-users";import "@pnp/sp/profiles";

export default class UserProfileService {
    private _LOG_SOURCE: string = "PersonalGreetingWebPart - UserProfileService";
    private _sp: SPFI;

    // Properties are stored in Key/Value pairs,
    // so parse into an object called userProperties
    private _profileProps = {};
    private _cibcPreferredName: string = '';

    public constructor(context?: WebPartContext) {
        this._sp = getSP();
    }

    // Get the current users SP profiles
    public async retrieveCurrentUserProfiles(): Promise<void> {
        try {
            const profile = await this._sp.profiles.myProperties();
            // const profile = await this._sp.profiles.myProperties.select("Title", "Email")();
            // console.log("Email: "+profile.Email);

            profile.UserProfileProperties.forEach((prop) => {                
                this._profileProps[prop.Key] = prop.Value;
                // console.log(prop.Key +": "+prop.Value);
                // Logger.write(`${this._LOG_SOURCE} (getCurrentUserProfiles) - ${LogLevel.Info}\n${prop.Key +": "+prop.Value} `, LogLevel.Info);            
            });            
        } catch (err) {
            Logger.write(`${this._LOG_SOURCE} (getCurrentUserProfiles) - ${LogLevel.Error}\n${err.message} ${JSON.stringify(err)}`, LogLevel.Error);
            throw Error(`${err.message}`);
        }
    }

    public async getCurrentUserPreferredName(): Promise<string> {
        try {    
            Object.keys(this._profileProps).forEach((key)=> {
                if (key === "CIBC-PreferredName") {
                    this._cibcPreferredName = this._profileProps[key];
                    // console.log("CIBC-PreferredName: " + this._cibcPreferredName);
                }
            });
            return this._cibcPreferredName;
        } catch (err) {
            Logger.write(`${this._LOG_SOURCE} (getCurrentUserProfiles) - ${LogLevel.Error}\n${err.message} ${JSON.stringify(err)}`, LogLevel.Error);
            throw Error(`${err.message}`);
        }
    }  
}
