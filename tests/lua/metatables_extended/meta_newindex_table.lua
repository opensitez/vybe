-- vybe-test: lua/metatables_extended/meta_newindex_table
-- origin: languages/lua/tests/lua/test_metatables_extended.rs

local __w1 = "42 nil"
local __i = 0

local parent = {}
local child = setmetatable({}, {__newindex = parent})
child.x = 42
do local __t = tostring(parent.x) .. "\t" .. tostring(tostring(rawget(child, "x"))); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
