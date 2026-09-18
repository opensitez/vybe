# vybe-test: powershell/xml_adapter_dom_navigation/comment_ignored_by_dot_navigation
$xml = [xml]'<ServerConfig>
  <!-- Main Database Configuration -->
  <Port>5432</Port>
  <!-- Connection Pool Settings -->
  <MaxConnections>100</MaxConnections>
</ServerConfig>'

# XML comments must be ignored by property dot navigation without polluting member resolution
$port = $xml.ServerConfig.Port
$maxConn = $xml.ServerConfig.MaxConnections

if ($port -ne "5432" -or $maxConn -ne "100") {
    Write-Host "FAIL: dot navigation failed in presence of comments, got port='$port', maxConn='$maxConn'"
    exit 1
}

Write-Host "PASS"
exit 0
