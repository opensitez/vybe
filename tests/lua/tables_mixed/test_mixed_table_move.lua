-- vybe-test: lua/tables_mixed/test_mixed_table_move
-- origin: languages/lua/tests/lua/test_tables_mixed.rs

local __w1 = "10 20 nil 2"
local __i = 0

local t={10, 20, a=1}; local t2={b=2}; table.move(t, 1, 2, 1, t2); do local __t = tostring(t2[1]..' '..t2[2]..' '..(t2.a or 'nil')..' '..t2.b); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
