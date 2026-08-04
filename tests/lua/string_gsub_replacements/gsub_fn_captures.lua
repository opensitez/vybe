-- vybe-test: lua/string_gsub_replacements/gsub_fn_captures
-- origin: languages/lua/tests/lua/test_string_gsub_replacements.rs

local __w1 = "a[1] b[2]"
local __i = 0

local r = string.gsub("a=1 b=2", "(%a+)=(%d+)", function(k, v) return k.."["..v.."]" end)
do local __t = tostring(r); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
