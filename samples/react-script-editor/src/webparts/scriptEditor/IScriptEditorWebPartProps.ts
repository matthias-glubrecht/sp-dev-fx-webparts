import type { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";

export interface IScriptEditorWebPartProps {
  script: string;
  useExternalScript: boolean;
  externalScript?: string;
  title: string;
  removePadding: boolean;
  spPageContextInfo: boolean;
  teamsContext: boolean;
  audiences: IPropertyFieldGroupOrPerson[];
  audienceCacheDuration: number;
}
