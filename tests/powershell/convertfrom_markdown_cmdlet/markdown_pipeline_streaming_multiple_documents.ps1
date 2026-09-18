# vybe-test: powershell/convertfrom_markdown_cmdlet/markdown_pipeline_streaming_multiple_documents
$docs = @("# First Document", "# Second Document")

# Piping multiple markdown strings into ConvertFrom-Markdown emits a MarkdownInfo for each input
$results = @($docs | ConvertFrom-Markdown)

if ($results.Count -ne 2) {
    Write-Host "FAIL: expected 2 converted documents, got $($results.Count)"
    exit 1
}

if ($results[0].Html -notmatch "First Document" -or $results[1].Html -notmatch "Second Document") {
    Write-Host "FAIL: pipeline converted markdown documents mismatch"
    exit 1
}

Write-Host "PASS"
exit 0
