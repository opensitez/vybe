lua_print! {
    test_iterator_stateless => {
        "local function iter(a, i)
    i = i + 1
    local v = a[i]
    if v then return i, v end
end
local function ipairs_custom(a)
    return iter, a, 0
end
local s = ''
for i, v in ipairs_custom({'a', 'b', 'c'}) do
    s = s .. v
end
print(s)",
        "abc"
    },
    test_iterator_stateful => {
        "local function iter(state)
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
print(s)",
        "xyz"
    },
    test_iterator_closure => {
        "local function values(t)
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
print(s)",
        "123"
    },
    test_iterator_fibonacci => {
        "local function fib(max)
    local a, b = 0, 1
    return function()
        if a > max then return nil end
        local curr = a
        a, b = b, a + b
        return curr
    end
end
local s = ''
for v in fib(10) do
    s = s .. v .. ' '
end
print(s)",
        "0 1 1 2 3 5 8 "
    },
    test_iterator_multi_return => {
        "local function pairs_custom(t)
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
print(s)",
        "a1b2c3"
    }
}
