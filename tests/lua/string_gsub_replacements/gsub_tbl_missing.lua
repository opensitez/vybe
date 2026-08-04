-- vybe-test: lua/string_gsub_replacements/gsub_tbl_missing
-- origin: languages/lua/tests/lua/test_string_gsub_replacements.rs

local __w1 = "HI world"
local __i = 0

local t = {hello="HI"}
local r = string.gsub("hello world", "%a+", t)
do local __t = tostring(r); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
