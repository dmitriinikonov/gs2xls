# GeoServer to Excel Report Script

This Python script connects to a GeoServer instance and generates an Excel report using data from the server.
The script creates the Workspaces, Stores, Layer Groups, Layers, Styles tabs in the Excel report, as well as a tab for each group. Where appropriate, the cells have hyperlinks - for example, from the Layer Groups tab you can follow the link to the tab of the corresponding group by clicking on its name in the Group Name field.

Cells are formatted with:
 - Dark blue font for 'N/A' values
 - Light blue background for cells containing descriptive metadata or status indicators

## Dependencies

Refer to `requirements.txt` for a complete list of dependencies.

## Security Note

**Important**: Ensure that your credentials are managed securely. Using environment variables is recommended to avoid hardcoding sensitive information.

## License

This project is licensed under the MIT License. See the `LICENSE` file for more details.
