-- vybe-test: lua/tables_hash/test_hash_add_during_iteration_error
-- origin: languages/lua/tests/lua/test_tables_hash.rs

local __w1 = "true 2"
local __i = 0

local t={a=1}; local ok = pcall(function() for k,v in pairs(t) do t.b=2 end end); do local __t = tostring(tostring(ok)..' '..t.b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
