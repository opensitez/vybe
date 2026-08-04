-- vybe-test: lua/metatables_index/test_index_rawget
-- origin: languages/lua/tests/lua/test_metatables_index.rs

local __w1 = "nil"
local __i = 0

local t=setmetatable({}, {__index={a=1}}); do local __t = tostring(rawget(t, 'a') or 'nil'); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
