### Generic Snippet

## Kill Specific Process for Current User

``` vb.net
PL.Where(function(x) x.ProcessName.ToLower = "mstsc" and x.SessionId = Process.GetCurrentProcess().SessionId).ToList()
```
