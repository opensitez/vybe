-- vybe-test: lua/metatables/rawget_reads_own_field_skipping_index_metamethod
-- origin: languages/lua/tests/lua/test_metatables.rs

local __w1 = "false,42"
local __i = 0

local fallback_used = false
local t = setmetatable({}, {
  __index = function() fallback_used = true; return 99 end
})
rawset(t, 'x', 42)
local v = rawget(t, 'x')
do local __t = tostring(tostring(fallback_used) .. ',' .. v); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
