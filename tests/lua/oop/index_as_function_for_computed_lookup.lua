-- vybe-test: lua/oop/index_as_function_for_computed_lookup
-- origin: languages/lua/tests/lua/test_oop.rs

local __w1 = "HELLOWORLD"
local __i = 0

local t = setmetatable({}, {
  __index = function(tbl, key)
    return key:upper()
  end
})
do local __t = tostring(t.hello .. t.world); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
