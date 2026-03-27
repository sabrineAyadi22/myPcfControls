/*
*This is auto generated from the ControlManifest.Input.xml file
*/

// Define IInputs and IOutputs Type. They should match with ControlManifest.
export interface IInputs {
    sellingPrice: ComponentFramework.PropertyTypes.DecimalNumberProperty;
    costPrice: ComponentFramework.PropertyTypes.DecimalNumberProperty;
    indiceName: ComponentFramework.PropertyTypes.StringProperty;
}
export interface IOutputs {
    resultValue?: number;
    debugText?: string;
    benefit?: string;
}
