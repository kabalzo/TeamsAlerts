<h1>Auto-deploy folders and files for a rudementary alerting system in Microsoft Teams</h1>
<ol>
  <li>Fill out AlertTemplate.csv with name, room, and messages</li>
    <ul>
      <li>Use an existing Teams team or make a new one.</li>
      <li>Create a new channel or use an existing one.</li>
      <li>Configure a webhook in a channel and copy the url to AlertTemplate.csv file.</li>
    </ul>
  <li>Run AutoConfigAlerts.ps1</li>
    <ul>
      <li>You will have to use Get-ExecutionPolicy and Set-ExecutionPolicy to run unsigned script.</li>
    </ul>
  <li>Create folder in Setup and copy/paste AutoConfig and AlertTemplate files into it</li>
    <ul>
      <li>If you filled out the .csv file correctly, everything will be configured automatically</li>
      <li>Creates new folders in root folder</li>
      <li>You can edit the .csv files and re-run AutoConfig.ps1</li>
      <li>Only configured to create three alerts with three icons. Change as needed</li>
      <li>Room numbers should be unique</li>
    </ul>
</ol>
