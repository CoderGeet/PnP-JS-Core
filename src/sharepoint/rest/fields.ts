"use strict";

import { Queryable, QueryableCollection, QueryableInstance } from "./queryable";
import { TypedHash } from "../../collections/collections";
import * as Util from "../../utils/util";
import * as Types from "./types";

/**
 * Describes a collection of Field objects
 * 
 */
export class Fields extends QueryableCollection {

    /**
     * Creates a new instance of the Fields class
     * 
     * @param baseUrl The url or Queryable which forms the parent of this fields collection
     */
    constructor(baseUrl: string | Queryable, path = "fields") {
        super(baseUrl, path);
    }

    /**
     * Gets a field from the collection by title
     * 
     * @param title The case-sensitive title of the field 
     */
    public getByTitle(title: string): Field {
        return new Field(this, `getByTitle('${title}')`);
    }

    /**
     * Gets a list from the collection by guid id
     * 
     * @param title The Id of the list  
     */
    public getById(id: string): Field {
        return new Field(this.toUrl().concat(`('${id}')`));
    }

    /**
     * Creates a field based on the specified schema
     */
    public createFieldAsXml(xml: string | Types.XmlSchemaFieldCreationInformation): Promise<FieldAddResult> {

        let info: Types.XmlSchemaFieldCreationInformation;
        if (typeof xml === "string") {
            info = { SchemaXml: xml };
        } else {
            info = xml as Types.XmlSchemaFieldCreationInformation;
        }

        let postBody = JSON.stringify({
            "parameters":
            Util.extend({
                "__metadata":
                {
                    "type": "SP.XmlSchemaFieldCreationInformation",
                },
            }, info),
        });

        let q = new Fields(this, "createfieldasxml");

        return q.post({ body: postBody }).then((data) => {
            return {
                data: data,
                field: this.getById(data.Id),
            };
        });
    }

    /**
     * Adds a new list to the collection
     * 
     * @param title The new field's title
     * @param fieldType The new field's type (ex: SP.FieldText)
     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
     */
    public add(title: string, fieldType: string, properties: TypedHash<string | number | boolean> = {}): Promise<FieldAddResult> {

        let postBody = JSON.stringify(Util.extend({
            "__metadata": { "type": fieldType },
            "Title": title,
        }, properties));

        return this.post({ body: postBody }).then((data) => {
            return {
                data: data,
                field: this.getById(data.Id),
            };
        });
    }

    /**
     * Adds a new SP.FieldText to the collection
     * 
     * @param title The field title
     * @param maxLength The maximum number of characters allowed in the value of the field.
     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
     */
    public addText(title: string, maxLength = 255, properties?: TypedHash<string | number | boolean>): Promise<FieldAddResult> {
        return this.add(title, "SP.FieldText", Util.extend({ FieldTypeKind: 2 }, properties));
    }

    /**
     * Adds a new SP.FieldCalculated to the collection
     * 
     * @param title The field title.
     * @param formula The formula for the field.
     * @param dateFormat The date and time format that is displayed in the field.
     * @param outputType Specifies the output format for the field. Represents a FieldType value.
     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
     */
    public addCalculated(
        title: string,
        formula: string,
        dateFormat: Types.DateTimeFieldFormatType,
        outputType: Types.FieldTypes = Types.FieldTypes.Text,
        properties?: TypedHash<string | number | boolean>): Promise<FieldAddResult> {
        return this.add(title, "SP.FieldCalculated", Util.extend({
            DateFormat: dateFormat,
            FieldTypeKind: 17,
            Formula: formula,
            OutputType: outputType,
        }, properties));
    }

    /**
     * Adds a new SP.FieldDateTime to the collection
     * 
     * @param title The field title
     * @param displayFormat The format of the date and time that is displayed in the field.
     * @param calendarType Specifies the calendar type of the field. 
     * @param properties Differ by type of field being created (see: https://msdn.microsoft.com/en-us/library/office/dn600182.aspx)
     */
    public addDateTime(
        title: string,
        displayFormat: Types.DateTimeFieldFormatType,
        calendarType: Types.CalendarType,
        friendlyDisplayFormat: ["0", "1", "2"],
        properties?: TypedHash<string | number | boolean>): Promise<FieldAddResult> {

        return this.add(title, "SP.FieldDateTime", Util.extend({
            DateTimeCalendarType: calendarType,
            DisplayFormat: displayFormat,
            FieldTypeKind: 4,
            FriendlyDisplayFormat: friendlyDisplayFormat,
        }, properties));
    }
}


/**
 * Describes a single of Field instance
 * 
 */
export class Field extends QueryableInstance {

    /**
     * Creates a new instance of the Field class
     * 
     * @param baseUrl The url or Queryable which forms the parent of this field instance
     */
    constructor(baseUrl: string | Queryable, path?: string) {
        super(baseUrl);
    }

    /**
     * Updates this field intance with the supplied properties 
     * 
     * @param properties A plain object hash of values to update for the list
     * @param fieldType The type value, required to update child field type properties
     */
    public update(properties: TypedHash<string | number | boolean>, fieldType = "SP.Field"): Promise<FieldUpdateResult> {

        let postBody = JSON.stringify(Util.extend({
            "__metadata": { "type": fieldType },
        }, properties));

        return this.post({
            body: postBody,
            headers: {
                "X-HTTP-Method": "MERGE",
            },
        }).then((data) => {
            return {
                data: data,
                field: this,
            };
        });
    }

    /**
     * Delete this fields
     * 
     */
    public delete(): Promise<void> {
        return this.post({
            headers: {
                "X-HTTP-Method": "DELETE",
            },
        });
    }

    /**
     * Sets the value of the ShowInDisplayForm property for this field.
     */
    public setShowInDisplayForm(show: boolean): Promise<void> {
        let q = new Field(this, `setshowindisplayform(${show})`);
        return q.post();
    }

    /**
     * Sets the value of the ShowInEditForm property for this field.
     */
    public setShowInEditForm(show: boolean): Promise<void> {
        let q = new Field(this, `setshowineditform(${show})`);
        return q.post();
    }

    /**
     * Sets the value of the ShowInNewForm property for this field.
     */
    public setShowInNewForm(show: boolean): Promise<void> {
        let q = new Field(this, `setshowinnewform(${show})`);
        return q.post();
    }
}

export interface FieldAddResult {
    data: any;
    field: Field;
}

export interface FieldUpdateResult {
    data: any;
    field: Field;
}
