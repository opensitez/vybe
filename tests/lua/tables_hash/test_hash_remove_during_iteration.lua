-- vybe-test: lua/tables_hash/test_hash_remove_during_iteration
-- origin: languages/lua/tests/lua/test_tables_hash.rs

local __w1 = "1 nil 3"
local __i = 0

local t={a=1, b=2, c=3}; for k,v in pairs(t) do if k=='b' then t[k]=nil end end; do local __t = tostring(t.a..' '..(t.b or 'nil')..' '..t.c); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
