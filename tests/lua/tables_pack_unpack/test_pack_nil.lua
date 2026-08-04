-- vybe-test: lua/tables_pack_unpack/test_pack_nil
-- origin: languages/lua/tests/lua/test_tables_pack_unpack.rs

local __w1 = "3 1 nil 3"
local __i = 0

local t = table.pack(1, nil, 3); do local __t = tostring(t.n..' '..t[1]..' '..tostring(t[2])..' '..t[3]); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
