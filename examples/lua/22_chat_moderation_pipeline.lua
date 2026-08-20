-- App 22: Chat Moderation Pipeline

local messages = {
  {user = "neo", text = "hello team"},
  {user = "max", text = "buy cheap stuff now"},
  {user = "sam", text = "this release is broken"},
  {user = "ivy", text = "visit badlink.example"},
}

local banned = {cheap = true, badlink = true}
local toxicity = {broken = 2, stupid = 3}

local function score_text(text)
  local score = 0
  for token in text:lower():gmatch("[%a_]+") do
    if toxicity[token] then score = score + toxicity[token] end
  end
  return score
end

local function has_banned(text)
  for token in text:lower():gmatch("[%a_]+") do
    if banned[token] then return true, token end
  end
  return false
end

local results = {}
for _, m in ipairs(messages) do
  local blocked, token = has_banned(m.text)
  local score = score_text(m.text)
  local action = "allow"
  if blocked then
    action = "block"
  elseif score >= 3 then
    action = "review"
  end
  results[#results + 1] = {user = m.user, action = action, reason = token or ("toxicity=" .. score)}
end

for _, r in ipairs(results) do
  print(string.format("user=%s action=%s reason=%s", r.user, r.action, r.reason))
end
