-- vybe-test: lua/loops_for_generic/generic_for_pcall_inside_body
-- origin: languages/lua/tests/lua/test_loops_for_generic.rs

local __w1 = "true,true,false,true"
local __i = 0

local results = {}
for i, v in ipairs({1, 2, 'bad', 4}) do
  local ok, n = pcall(function() return v + 0 end)
  results[i] = tostring(ok)
end
do local __t = tostring(table.concat(results, ',')); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
