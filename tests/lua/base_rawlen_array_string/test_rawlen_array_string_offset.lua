-- vybe-test: lua/base_rawlen_array_string/test_rawlen_array_string_offset
-- origin: languages/lua/tests/lua/test_base_rawlen_array_string.rs

local __w1 = "true"
local __i = 0

local t = { [1] = "x", [2] = "y", [4] = "z"}; do local __t = tostring(type(rawlen(t)) == "number"); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
