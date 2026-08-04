-- vybe-test: lua/metatables_extended/meta_index_function
-- origin: languages/lua/tests/lua/test_metatables_extended.rs

local __w1 = "hello!"
local __i = 0

local t = setmetatable({}, {__index = function(tbl, key) return key .. "!" end})
do local __t = tostring(t.hello); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
