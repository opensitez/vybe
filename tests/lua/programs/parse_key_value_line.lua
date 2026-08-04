-- vybe-test: lua/programs/parse_key_value_line
-- origin: languages/lua/tests/lua/test_programs.rs

local __w1 = "name:vybe"
local __i = 0

local line = "name=vybe"
local k, v = string.match(line, "(%w+)=(%w+)")
do local __t = tostring(k .. ":" .. v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
