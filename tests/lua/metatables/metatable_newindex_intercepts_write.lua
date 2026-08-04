-- vybe-test: lua/metatables/metatable_newindex_intercepts_write
-- origin: languages/lua/tests/lua/test_metatables.rs

local __w1 = "x=10,y=20"
local __i = 0

local log = {}
local t = setmetatable({}, {
  __newindex = function(tbl, key, val)
    log[#log+1] = key .. '=' .. tostring(val)
    rawset(tbl, key, val)
  end
})
t.x = 10
t.y = 20
do local __t = tostring(table.concat(log, ',')); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
