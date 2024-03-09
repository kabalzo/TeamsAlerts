<h1>Auto-deploy folders and files for a rudementary alerting system in Microsoft Teams</h1>
<h3>I do not consider myself a professional coder. I am a hobbyist trying to build experience where I can.</h3>
<h3>Sorry for spaghetti code. I welcome suggestions and improvements.</h3>
<br></br>
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
  <li>When prompted, enter name of the directory you want files and sub-directories to be placed in</li>
    <ul>
      <li>This name will become a part of the .ps1 files also.</li>
      <li>Example: Running AutoConfigAlerts.ps1 from with C:\Test and specifying XYZ will create C:\Test\XYZ directory.</li>
      <li>Confirm you entered the correct name.</li>
    </ul>
</ol>
