-- vybe-test: lua/string_patterns_advanced/test_string_gsub_with_function_multiple_args
-- origin: languages/lua/tests/lua/test_string_patterns_advanced.rs

local __w1 = "x20, y40"
local __i = 0

do local __t = tostring((string.gsub('x=10, y=20', '(%w+)=(%d+)', function(k,v) return k..tonumber(v)*2 end))); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
