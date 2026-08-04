-- vybe-test: lua/tables_custom_iterators/test_iterator_stateful
-- origin: languages/lua/tests/lua/test_tables_custom_iterators.rs

local __w1 = "xyz"
local __i = 0

local function iter(state)
    state.i = state.i + 1
    local v = state.a[state.i]
    if v then return state.i, v end
end
local function ipairs_custom(a)
    return iter, {a = a, i = 0}
end
local s = ''
for i, v in ipairs_custom({'x', 'y', 'z'}) do
    s = s .. v
end
do local __t = tostring(s); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
