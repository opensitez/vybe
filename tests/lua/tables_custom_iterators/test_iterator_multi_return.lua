-- vybe-test: lua/tables_custom_iterators/test_iterator_multi_return
-- origin: languages/lua/tests/lua/test_tables_custom_iterators.rs

local __w1 = "a1b2c3"
local __i = 0

local function pairs_custom(t)
    local keys = {}
    for k in pairs(t) do table.insert(keys, k) end
    table.sort(keys)
    local i = 0
    return function()
        i = i + 1
        local k = keys[i]
        if k then return k, t[k] end
    end
end
local s = ''
for k, v in pairs_custom({b = 2, a = 1, c = 3}) do
    s = s .. k .. v
end
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
