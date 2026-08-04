-- vybe-test: lua/metatables_index_chains/index_func_args
-- origin: languages/lua/tests/lua/test_metatables_index_chains.rs

local __w1 = "table,foo"
local __i = 0

local log = nil
local t = setmetatable({}, {
  __index = function(tbl, k)
    log = type(tbl) .. "," .. k
    return 0
  end
})
_ = t.foo
do local __t = tostring(log); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
