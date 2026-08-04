-- vybe-test: lua/metatables_extended/meta_newindex_function
-- origin: languages/lua/tests/lua/test_metatables_extended.rs

local __w1 = "score=100"
local __i = 0

local log = nil
local t = setmetatable({}, {__newindex = function(tbl, k, v) log = k .. "=" .. v end})
t.score = 100
do local __t = tostring(log); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
