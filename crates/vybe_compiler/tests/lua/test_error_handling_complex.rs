lua_print! {
    test_xpcall_message_handler_modifies_error => {
        "local function handler(err)
    return 'caught: ' .. tostring(err)
end
local ok, res = xpcall(function() error('boom') end, handler)
print(res)",
        "caught: boom"
    },
    test_xpcall_handler_throws_error => {
        "local function handler(err)
    error('double fault')
end
local ok, res = xpcall(function() error('boom') end, handler)
print(tostring(string.find(res, 'error in error handling') ~= nil))",
        "true"
    },
    test_xpcall_args => {
        "local function handler(err)
    return 'caught'
end
local ok, res = xpcall(function(a, b) return a + b end, handler, 10, 20)
print(res)",
        "30"
    },
    test_pcall_multiple_returns => {
        "local function multi() return 1, 2, 3 end
local a,b,c,d = pcall(multi)
print(tostring(a) .. ' ' .. b .. ' ' .. c .. ' ' .. d)",
        "true 1 2 3"
    },
    test_pcall_varargs => {
        "local function sum(...)
    local s = 0
    for _, v in ipairs({...}) do s = s + v end
    return s
end
local ok, res = pcall(sum, 1, 2, 3, 4)
print(res)",
        "10"
    }
}
