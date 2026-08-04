-- vybe-test: lua/string_gsub_ext/test_gsub_function_returns_nil
-- origin: languages/lua/tests/lua/test_string_gsub_ext.rs

local __w1 = "a b c"
local __i = 0

local s = string.gsub('a b c', 'b', function() return nil end); do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
