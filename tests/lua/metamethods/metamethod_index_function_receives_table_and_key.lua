-- vybe-test: lua/metamethods/metamethod_index_function_receives_table_and_key
-- origin: languages/lua/tests/lua/test_metamethods.rs

local __w1 = "k:foo"
local __i = 0

local t = setmetatable({}, {__index = function(tbl, k) return "k:" .. k end})
do local __t = tostring(t.foo); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
