-- vybe-test: lua/debug_hooks/test_debug_sethook_count
-- origin: languages/lua/tests/lua/test_debug_hooks.rs

local __w1 = "true"
local __i = 0

local c=0; debug.sethook(function() c=c+1 end, '', 10); for i=1,20 do end; debug.sethook(); do local __t = tostring(tostring(c > 0)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
