lua_print! {
    test_proxy_read_only => {
        "local function ReadOnly(t)
    local proxy = {}
    local mt = {
        __index = t,
        __newindex = function(t, k, v)
            error('attempt to update a read-only table')
        end
    }
    setmetatable(proxy, mt)
    return proxy
end
local data = {a = 1}
local ro_data = ReadOnly(data)
local ok, err = pcall(function() ro_data.a = 2 end)
print(ro_data.a .. ' ' .. tostring(string.find(err, 'read-only') ~= nil))",
        "1 true"
    },
    test_proxy_strict_mode => {
        "local function Strict(t)
    local mt = {
        __index = function(t, k)
            error('attempt to read undeclared variable ' .. k)
        end,
        __newindex = function(t, k, v)
            error('attempt to write to undeclared variable ' .. k)
        end
    }
    setmetatable(t, mt)
    return t
end
local strict_tbl = Strict({})
local ok_read, err_read = pcall(function() return strict_tbl.undeclared end)
local ok_write, err_write = pcall(function() strict_tbl.undeclared = 1 end)
print(tostring(string.find(err_read, 'undeclared') ~= nil) .. ' ' .. tostring(string.find(err_write, 'undeclared') ~= nil))",
        "true true"
    },
    test_proxy_default_value => {
        "local function DefaultTable(default)
    local mt = {
        __index = function(t, k)
            return default
        end
    }
    return setmetatable({}, mt)
end
local counts = DefaultTable(0)
counts['a'] = counts['a'] + 1
print(counts['a'] .. ' ' .. counts['b'])",
        "1 0"
    },
    test_proxy_tracking_access => {
        "local function Track(t)
    local proxy = {}
    local accesses = 0
    local mt = {
        __index = function(_, k)
            accesses = accesses + 1
            return t[k]
        end,
        __newindex = function(_, k, v)
            t[k] = v
        end
    }
    setmetatable(proxy, mt)
    return proxy, function() return accesses end
end
local p, get_accesses = Track({x = 10, y = 20})
local a = p.x + p.y
p.z = 30
print(a .. ' ' .. p.z .. ' ' .. get_accesses())",
        "30 30 3"
    }
}
