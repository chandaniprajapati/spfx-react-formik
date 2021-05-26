import * as React from 'react';
import styles from './ReactFormik.module.scss';
import { IReactFormikProps } from './IReactFormikProps';
import { IReactFormikState } from './IReactFormikState';
import { SPService } from '../../shared/service/SPService';
import { TextField, MaskedTextField } from '@fluentui/react/lib/TextField';
import { Stack, IStackProps, IStackStyles } from '@fluentui/react/lib/Stack';
import { Form, Formik, Field, FormikProps } from 'formik';
import { Label } from 'office-ui-fabric-react/lib/Label';
import * as yup from 'yup';
import { PeoplePicker, PrincipalType } from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { DatePicker, Dropdown, mergeStyleSets, PrimaryButton, IIconProps, IDropdownOption, Grid, VirtualizedComboBox, IComboBoxOption } from 'office-ui-fabric-react';
import { sp } from '@pnp/sp';
import { Dialog } from '@microsoft/sp-dialog';

const stackTokens = { childrenGap: 50 };
const iconProps = { iconName: 'Calendar' };
const stackStyles: Partial<IStackStyles> = { root: { width: 650 } };
const columnProps: Partial<IStackProps> = {
  tokens: { childrenGap: 15 },
  styles: { root: { width: 300 } },
};
const controlClass = mergeStyleSets({
  control: {
    margin: '0 0 15px 0',
    maxWidth: '300px',
  },
});

export default class ReactFormik extends React.Component<IReactFormikProps, IReactFormikState> {

  private cancelIcon: IIconProps = { iconName: 'Cancel' };
  private saveIcon: IIconProps = { iconName: 'Save' };
  private _services: SPService = null;

  constructor(props: Readonly<IReactFormikProps>) {
    super(props);
    this.state = {
      startDate: null,
      endDate: null
    }
    sp.setup({
      spfxContext: this.props.context
    });

    this._services = new SPService(this.props.siteUrl);
    this.createRecord = this.createRecord.bind(this);
  }

  private getFieldProps = (formik: FormikProps<any>, field: string) => {
    return { ...formik.getFieldProps(field), errorMessage: formik.errors[field] as string }
  }

  public async createRecord(record: any) {
    let item = await this._services.createTask("Tasks", {
      Title: record.name,
      TaskDetails: record.details,
      StartDate: record.startDate,
      EndDate: new Date(record.endDate),
      ProjectName: record.projectName,
    }).then((data) => {
      Dialog.alert("Record inseterd successfully !!!");
      return data;
    }).catch((err) => {
      console.error(err);
      throw err;
    });
  }

  public render(): React.ReactElement<IReactFormikProps> {
    const validate = yup.object().shape({
      name: yup.string().required('Task name is required'),
      details: yup.string()
        .min('15', 'Minimum required 15 characters')
        .required('Task detail is required'),
      projectName: yup.string().required('Please select a project'),
      startDate: yup.date().required('Please select a start date').nullable(),
      endDate: yup.date().required('Please select a end date').nullable()
    })

    return (
      <Formik initialValues={{
        name: '',
        details: '',
        projectName: '',
        startDate: null,
        endDate: null
      }}
        validationSchema={validate}
        onSubmit={(values, helpers) => {
          console.log('SUCCESS!! :-)\n\n' + JSON.stringify(values, null, 4));
          this.createRecord(values).then(response => {
            helpers.resetForm()
          });
        }}>
        { formik => (
          <div className={styles.reactFormik}>
            <Stack>
              <Label className={styles.lblForm}>Current User</Label>
              <PeoplePicker
                context={this.props.context as any}
                personSelectionLimit={1}
                showtooltip={true}
                showHiddenInUI={false}
                principalTypes={[PrincipalType.User]}
                ensureUser={true}
                disabled={true}
                defaultSelectedUsers={[this.props.context.pageContext.user.email as any]}
              />

              <Label className={styles.lblForm}>Task Name</Label>
              <TextField
                {...this.getFieldProps(formik, 'name')} />

              <Label className={styles.lblForm}>Project Name</Label>
              <Dropdown
                options={
                  [
                    { key: 'Project1', text: 'Project1' },
                    { key: 'Project2', text: 'Project2' },
                    { key: 'Project3', text: 'Project3' },
                  ]
                }
                {...this.getFieldProps(formik, 'projectName')}
                onChange={(event, option) => { formik.setFieldValue('projectName', option.key.toString()) }}
              />

              <Label className={styles.lblForm}>Start Date</Label>
              <DatePicker
                className={controlClass.control}
                id="startDate"
                value={formik.values.startDate}
                textField={{ ...this.getFieldProps(formik, 'startDate') }}
                onSelectDate={(date) => formik.setFieldValue('startDate', date)}
              />

              <Label className={styles.lblForm}>End Date</Label>
              <DatePicker
                className={controlClass.control}
                id="endDate"
                value={formik.values.endDate}
                textField={{ ...this.getFieldProps(formik, 'endDate') }}
                onSelectDate={(date) => formik.setFieldValue('endDate', date)}
              />

              <Label className={styles.lblForm}>Task Details</Label>
              <TextField
                multiline
                rows={6}
                {...this.getFieldProps(formik, 'details')} />

            </Stack>
            <PrimaryButton
              type="submit"
              text="Save"
              iconProps={this.saveIcon}
              className={styles.btnsForm}
              onClick={formik.handleSubmit as any}
            />
            <PrimaryButton
              text="Cancel"
              iconProps={this.cancelIcon}
              className={styles.btnsForm}
              onClick={formik.handleReset}
            />
          </div>
        )
        }
      </Formik >
    );
  }
}