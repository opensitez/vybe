-- vybe-test: lua/string_gsub_ext/test_gsub_function_multiple_captures
-- origin: languages/lua/tests/lua/test_string_gsub_ext.rs

local __w1 = "x10"
local __i = 0

local s = string.gsub('x=10', '(%a)=(%d+)', function(k,v) return k..v end); do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
