-- vybe-test: lua/metatables/__index_metamethod_on_metatable_itself
-- origin: languages/lua/tests/lua/test_metatables.rs

local __w1 = "hello!"
local __i = 0

local mt = {__index = function(_, k) return k .. "!" end}
local t = setmetatable({}, mt)
do local __t = tostring(t.hello); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
