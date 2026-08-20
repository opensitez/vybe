-- 14_utf8_text_processing.lua
-- Demonstrates utf8 library and codepoint-safe iteration.

local text = "Lua cafe: cafe\u{301} | Rocket: \u{1F680}"

print("raw bytes:", #text)
print("utf8 length:", utf8.len(text))

for pos, code in utf8.codes(text) do
  print(string.format("pos=%d codepoint=U+%X", pos, code))
end

local function reverse_utf8(s)
  local chars = {}
  for _, code in utf8.codes(s) do
    chars[#chars + 1] = utf8.char(code)
  end
  for i = 1, math.floor(#chars / 2) do
    chars[i], chars[#chars - i + 1] = chars[#chars - i + 1], chars[i]
  end
  return table.concat(chars)
end

print("reversed:", reverse_utf8(text))
