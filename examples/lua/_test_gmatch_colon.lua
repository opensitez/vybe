local function has_banned(text)
  for token in text:lower():gmatch("[%a_]+") do
    if token == "cheap" then return true, token end
  end
  return false
end
local blocked, token = has_banned("buy cheap stuff")
print(tostring(blocked), token or "")
