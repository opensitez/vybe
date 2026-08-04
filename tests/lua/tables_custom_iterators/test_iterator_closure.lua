-- vybe-test: lua/tables_custom_iterators/test_iterator_closure
-- origin: languages/lua/tests/lua/test_tables_custom_iterators.rs

local __w1 = "123"
local __i = 0

local function values(t)
    local i = 0
    return function()
        i = i + 1
        return t[i]
    end
end
local s = ''
for v in values({'1', '2', '3'}) do
    s = s .. v
end
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
