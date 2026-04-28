/*
*This is auto generated from the ControlManifest.Input.xml file
*/

// Define IInputs and IOutputs Type. They should match with ControlManifest.
export interface IInputs {
    indicesJson: ComponentFramework.PropertyTypes.StringProperty;
    DefaultValue: ComponentFramework.PropertyTypes.StringProperty;
    GapPercentThreshold: ComponentFramework.PropertyTypes.DecimalNumberProperty;
    ShowValidation: ComponentFramework.PropertyTypes.TwoOptionsProperty;
}
export interface IOutputs {
    OutputData?: string;
    HasValidationError?: boolean;
    ValidationErrorRows?: string;
    ShowCommentPopup?: boolean;
    CommentPopupData?: string;
    ValidationStatusSummary?: string;
}
