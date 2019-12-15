import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IDateTimeFieldValue } from "@pnp/spfx-property-controls/lib/PropertyFieldDateTimePicker";

export interface IUserSecretSantaProps {
  description: string;
  context:WebPartContext;
  propsdatetime:IDateTimeFieldValue;
}
