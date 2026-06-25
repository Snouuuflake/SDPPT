Write-Host "next slide";

$ppt = New-Object -ComObject PowerPoint.Application;
if ($ppt.SlideShowWindows.Count -ne 0) {
  $window = $ppt.SlideShowWindows.Item(1);
  $window.View.Previous();
} else {
  Write-Host "No slideshow windows";
}
