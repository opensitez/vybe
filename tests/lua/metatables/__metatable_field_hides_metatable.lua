-- vybe-test: lua/metatables/__metatable_field_hides_metatable
-- origin: languages/lua/tests/lua/test_metatables.rs

local __w1 = "true"
local __i = 0

local hidden = {}
local t = setmetatable({}, {__metatable=hidden})
do local __t = tostring(getmetatable(t) == hidden); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
