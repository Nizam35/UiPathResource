## UiPath Nuget point to Load the UiPath Packages on Prod

```
https://uipathpackages.myget.org/F/packages/api/v3/index.json
```

## Basic Nuget.Config Strucure
```
<?xml version="1.0" encoding="utf-8"?>
<configuration>
  <packageSources>
    <add key="nuget.org" value="https://api.nuget.org/v3/index.json" />
    <add key="Go" value="https://gallery.uipath.com/api/v2" />
    <add key="Official" value="https://www.myget.org/F/workflow" />
    <add key="UiPath Package" value="https://uipathpackages.myget.org/F/packages/api/v3/index.json" />
  </packageSources>
  <disabledPackageSources>
    <add key="nuget.org" value="true" />
    <add key="Go" value="true" />
  </disabledPackageSources>
</configuration>
```
