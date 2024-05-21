
$page = 1
$maxPageSize = 100

$url = $args[0] + "/api/publica/v2/chamados?size=" + $maxPageSize
$username = $args[1]
$password = $args[2]

function CallApi($url) {
    $base64AuthInfo = [Convert]::ToBase64String([Text.Encoding]::ASCII.GetBytes("${username}:${password}"))
    $headers = @{
        Authorization = "Basic ${base64AuthInfo}"
    }
    
    return Invoke-RestMethod -Uri $url -Headers $headers
}

$response = CallApi "$url&page=$page"
$responseSize = $response.Count

while ($responseSize -eq $maxPageSize) {
    $page++
    $newAPICall = CallApi "$url&page=$page"

    $response += $newAPICall
    $responseSize = $newAPICall.Count
}

$tickets = "-"
foreach ($item in $response) {
    if ($item.agente -eq $null) {
        $ticketNumber = $item.ticketKey
        $tickets = $tickets + "," + $ticketNumber + "ZGXT" + $item.titulo + "ZGXT" + $item.reporter.nome + "ZGXT" + $item.url
    }
}

Write-Output $tickets | Out-File -FilePath $args[3]
