-- vybe-test: lua/raw_access/rawset_writes_without_newindex
-- origin: languages/lua/tests/lua/test_raw_access.rs

local __w1 = "7"
local __i = 0

local store = {}
local t = setmetatable({}, {__newindex = function(_, k, v) store[k] = v end})
rawset(t, "k", 7)
do local __t = tostring(rawget(t, "k")); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
