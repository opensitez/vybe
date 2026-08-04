-- vybe-test: lua/metatables/rawset_does_not_invoke_newindex_when_writing
-- origin: languages/lua/tests/lua/test_metatables.rs

local __w1 = "false,val"
local __i = 0

local blocked = false
local t = setmetatable({}, {
  __newindex = function() blocked = true end
})
rawset(t, 'key', 'val')
do local __t = tostring(tostring(blocked) .. ',' .. t.key); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
