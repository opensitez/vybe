-- vybe-test: lua/goto/goto_exits_nested_ipairs_iteration
-- origin: languages/lua/tests/lua/test_goto.rs

local __w1 = "4"
local __i = 0

local found = nil
for _, row in ipairs({{1,2},{3,4},{5,6}}) do
  for _, v in ipairs(row) do
    if v == 4 then found = v; goto done end
  end
end
::done::
do local __t = tostring(found); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
