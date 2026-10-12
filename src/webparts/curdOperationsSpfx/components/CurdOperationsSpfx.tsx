import * as React from 'react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import {
  Box,
  Paper,
  Typography,
  Table,
  TableBody,
  TableCell,
  TableContainer,
  TableHead,
  TableRow,
  Button,
  TextField,
  Stack,
  Alert
} from '@mui/material';

import type { ICurdOperationsSpfxProps } from './ICurdOperationsSpfxProps';

export interface IEmployeeItem {
  Id?: number;
  Title: string;
  HireDate: string;
  JobDescription: string;
}

interface ISharePointListItem {
  Id?: number;
  [fieldName: string]: unknown;
}

interface ISharePointField {
  Title?: string;
  InternalName?: string;
  EntityPropertyName?: string;
  Hidden?: boolean;
}

interface IEmployeeFieldNames {
  employeeName: string;
  hireDate: string;
  jobDescription: string;
}

export interface ICurdOperationsSpfxState {
  items: IEmployeeItem[];
  selectedId: number | null;
  employeeName: string;
  hireDate: string;
  jobDescription: string;
  statusMessage: string;
  isError: boolean;
  isCreating: boolean;
}

export default class CurdOperationsSpfx extends React.Component<
  ICurdOperationsSpfxProps,
  ICurdOperationsSpfxState
> {
  private employeeFieldNames?: IEmployeeFieldNames;
  private employeeFieldListName?: string;

  constructor(props: ICurdOperationsSpfxProps) {
    super(props);
    this.state = {
      items: [],
      selectedId: null,
      employeeName: '',
      hireDate: '',
      jobDescription: '',
      statusMessage: '',
      isError: false,
      isCreating: false
    };
  }

  public componentDidMount(): void {
    void this.fetchEmployees();
  }

  private getListUrl(): string {
    return `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/items`;
  }

  private getEmployeeFieldNames = async (): Promise<IEmployeeFieldNames> => {
    if (this.employeeFieldNames && this.employeeFieldListName === this.props.listName) {
      return this.employeeFieldNames;
    }

    const fieldsUrl = `${this.props.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('${this.props.listName}')/fields?$select=Title,InternalName,EntityPropertyName,Hidden`;
    const response: SPHttpClientResponse = await this.props.context.spHttpClient.get(
      fieldsUrl,
      SPHttpClient.configurations.v1
    );
    if (!response.ok) {
      throw new Error(`Could not read SharePoint list columns (HTTP ${response.status}).`);
    }

    const data = await response.json();
    const fields: ISharePointField[] = (data.value || []).filter((field: ISharePointField) => !field.Hidden);
    const normalized = (value: string | undefined): string =>
      (value || '').replace(/[^a-z0-9]/gi, '').toLowerCase();
    const findField = (aliases: string[]): string | undefined => {
      const field = fields.find((candidate) =>
        aliases.indexOf(normalized(candidate.Title)) !== -1 ||
        aliases.indexOf(normalized(candidate.InternalName)) !== -1
      );
      return field ? field.EntityPropertyName || field.InternalName : undefined;
    };

    const employeeFieldNames: IEmployeeFieldNames = {
      employeeName: findField(['employeename', 'title']) || '',
      hireDate: findField(['hiredate']) || '',
      jobDescription: findField(['jobdescription']) || ''
    };
    const missingFields = [
      !employeeFieldNames.employeeName && 'Employee Name',
      !employeeFieldNames.hireDate && 'Hire Date',
      !employeeFieldNames.jobDescription && 'Job Description'
    ].filter((field): field is string => Boolean(field));
    if (missingFields.length > 0) {
      throw new Error(`Could not find these columns in the "${this.props.listName}" list: ${missingFields.join(', ')}.`);
    }

    this.employeeFieldNames = employeeFieldNames;
    this.employeeFieldListName = this.props.listName;
    return employeeFieldNames;
  };

  private fetchEmployees = async (): Promise<void> => {
    try {
      const fieldNames = await this.getEmployeeFieldNames();
      const url = `${this.getListUrl()}?$select=Id,${fieldNames.employeeName},${fieldNames.hireDate},${fieldNames.jobDescription}&$orderby=Id desc`;
      const response: SPHttpClientResponse = await this.props.context.spHttpClient.get(
        url,
        SPHttpClient.configurations.v1
      );

      if (response.ok) {
        const data = await response.json();
        const items: IEmployeeItem[] = (data.value || []).map((item: ISharePointListItem) => {
          const hireDateValue = item[fieldNames.hireDate];
          const employeeNameValue = item[fieldNames.employeeName];
          const jobDescriptionValue = item[fieldNames.jobDescription];
          return {
            Id: item.Id,
            Title: typeof employeeNameValue === 'string' ? employeeNameValue : '',
            HireDate: typeof hireDateValue === 'string' ? hireDateValue.split('T')[0] : '',
            JobDescription: typeof jobDescriptionValue === 'string' ? jobDescriptionValue : ''
          };
        });
        this.setState({ items });
      } else {
        this.setState({
          statusMessage: `Failed to read list items (HTTP ${response.status}).`,
          isError: true
        });
      }
    } catch (error) {
      const details = error instanceof Error ? ` ${error.message}` : '';
      this.setState({ statusMessage: `Failed to read list items.${details}`, isError: true });
    }
  };

  private handleCreate = async (): Promise<void> => {
    const { employeeName, hireDate, jobDescription } = this.state;
    if (!employeeName.trim()) {
      this.setState({ statusMessage: 'Employee Name is required.', isError: true });
      return;
    }

    this.setState({ isCreating: true, statusMessage: '', isError: false });
    try {
      const fieldNames = await this.getEmployeeFieldNames();
      const response: SPHttpClientResponse = await this.props.context.spHttpClient.post(
        this.getListUrl(),
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'odata-version': ''
          },
          body: JSON.stringify({
            [fieldNames.employeeName]: employeeName,
            [fieldNames.hireDate]: hireDate ? new Date(hireDate).toISOString() : null,
            [fieldNames.jobDescription]: jobDescription
          })
        }
      );

      if (response.ok) {
        this.setState({ statusMessage: 'Employee created successfully!', isError: false });
        this.clearForm();
        void this.fetchEmployees();
      } else {
        const responseBody = await response.text();
        const details = responseBody.trim();
        this.setState({
          statusMessage: `Create failed (HTTP ${response.status}${response.statusText ? ` ${response.statusText}` : ''})${details ? `: ${details}` : '.'}`,
          isError: true
        });
      }
    } catch (error) {
      const details = error instanceof Error ? `: ${error.message}` : '';
      this.setState({ statusMessage: `Error creating employee${details}.`, isError: true });
    } finally {
      this.setState({ isCreating: false });
    }
  };

  private handleUpdate = async (): Promise<void> => {
    const { selectedId, employeeName, hireDate, jobDescription } = this.state;
    if (!selectedId) {
      this.setState({ statusMessage: 'Please select a row to update.', isError: true });
      return;
    }

    try {
      const fieldNames = await this.getEmployeeFieldNames();
      const url = `${this.getListUrl()}(${selectedId})`;
      const response: SPHttpClientResponse = await this.props.context.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'X-HTTP-Method': 'MERGE',
            'IF-MATCH': '*'
          },
          body: JSON.stringify({
            [fieldNames.employeeName]: employeeName,
            [fieldNames.hireDate]: hireDate ? new Date(hireDate).toISOString() : null,
            [fieldNames.jobDescription]: jobDescription
          })
        }
      );

      if (response.ok) {
        this.setState({ statusMessage: 'Employee updated successfully!', isError: false });
        this.clearForm();
        void this.fetchEmployees();
      }
    } catch {
      this.setState({ statusMessage: 'Error updating employee.', isError: true });
    }
  };

  private handleDelete = async (): Promise<void> => {
    const { selectedId } = this.state;
    if (!selectedId) {
      this.setState({ statusMessage: 'Please select a row to delete.', isError: true });
      return;
    }

    try {
      const url = `${this.getListUrl()}(${selectedId})`;
      const response: SPHttpClientResponse = await this.props.context.spHttpClient.post(
        url,
        SPHttpClient.configurations.v1,
        {
          headers: {
            'Accept': 'application/json;odata=nometadata',
            'Content-type': 'application/json;odata=nometadata',
            'X-HTTP-Method': 'DELETE',
            'IF-MATCH': '*'
          }
        }
      );

      if (response.ok) {
        this.setState({ statusMessage: 'Employee deleted successfully!', isError: false });
        this.clearForm();
        void this.fetchEmployees();
      }
    } catch {
      this.setState({ statusMessage: 'Error deleting employee.', isError: true });
    }
  };

  private selectRow = (item: IEmployeeItem): void => {
    this.setState({
      selectedId: item.Id || null,
      employeeName: item.Title,
      hireDate: item.HireDate,
      jobDescription: item.JobDescription
    });
  };

  private clearForm = (): void => {
    this.setState({
      selectedId: null,
      employeeName: '',
      hireDate: '',
      jobDescription: ''
    });
  };

  public render(): React.ReactElement<ICurdOperationsSpfxProps> {
    const { items, selectedId, employeeName, hireDate, jobDescription, statusMessage, isError, isCreating } = this.state;

    return (
      <Box sx={{ maxWidth: 850, mx: 'auto', p: 2 }}>
        <Paper elevation={2} sx={{ p: 3, border: '1px solid #c8c8c8', borderRadius: 1 }}>
          <Typography variant="h5" sx={{ fontWeight: 'bold', mb: 2 }}>
            {this.props.description}
          </Typography>

          {statusMessage && (
            <Alert severity={isError ? 'error' : 'success'} sx={{ mb: 2 }}>
              {statusMessage}
            </Alert>
          )}

          <TableContainer component={Paper} variant="outlined" sx={{ mb: 2.5 }}>
            <Table size="small">
              <TableHead sx={{ backgroundColor: '#fafafa' }}>
                <TableRow>
                  <TableCell sx={{ fontWeight: 'bold' }}>Employee Name</TableCell>
                  <TableCell sx={{ fontWeight: 'bold' }}>Hire Date</TableCell>
                  <TableCell sx={{ fontWeight: 'bold' }}>Job Description</TableCell>
                </TableRow>
              </TableHead>
              <TableBody>
                {items.length === 0 ? (
                  <TableRow>
                    <TableCell colSpan={3} align="center">No employees found.</TableCell>
                  </TableRow>
                ) : (
                  items.map((item) => (
                    <TableRow
                      key={item.Id}
                      hover
                      onClick={() => this.selectRow(item)}
                      sx={{
                        cursor: 'pointer',
                        backgroundColor: selectedId === item.Id ? '#e0e0e0 !important' : 'inherit'
                      }}
                    >
                      <TableCell>{item.Title}</TableCell>
                      <TableCell>{item.HireDate}</TableCell>
                      <TableCell>{item.JobDescription}</TableCell>
                    </TableRow>
                  ))
                )}
              </TableBody>
            </Table>
          </TableContainer>

          <Stack direction="row" spacing={1.5} sx={{ mb: 3 }}>
            <Button
              variant="contained"
              sx={{ backgroundColor: '#2e7d32', '&:hover': { backgroundColor: '#1b5e20' } }}
              onClick={this.handleCreate}
              disabled={isCreating}
            >
              {isCreating ? 'Creating...' : 'Create'}
            </Button>
            <Button
              variant="contained"
              sx={{ backgroundColor: '#2e7d32', '&:hover': { backgroundColor: '#1b5e20' } }}
              onClick={this.handleUpdate}
              disabled={!selectedId}
            >
              Update
            </Button>
            <Button
              variant="contained"
              sx={{ backgroundColor: '#2e7d32', '&:hover': { backgroundColor: '#1b5e20' } }}
              onClick={this.handleDelete}
              disabled={!selectedId}
            >
              Delete
            </Button>
            <Button variant="outlined" color="inherit" onClick={this.clearForm}>
              Clear
            </Button>
          </Stack>

          <Box component="form" noValidate>
            <TextField
              label="Employee Name"
              fullWidth
              size="small"
              margin="dense"
              value={employeeName}
              onChange={(e) => this.setState({ employeeName: e.target.value })}
            />
            <TextField
              label="Hire Date"
              type="date"
              fullWidth
              size="small"
              margin="dense"
              InputLabelProps={{ shrink: true }}
              value={hireDate}
              onChange={(e) => this.setState({ hireDate: e.target.value })}
            />
            <TextField
              label="Job Description"
              fullWidth
              multiline
              rows={3}
              margin="dense"
              value={jobDescription}
              onChange={(e) => this.setState({ jobDescription: e.target.value })}
            />
          </Box>
        </Paper>
      </Box>
    );
  }
}
