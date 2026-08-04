-- vybe-test: lua/metatables_proxy_tables/test_proxy_read_only
-- origin: languages/lua/tests/lua/test_metatables_proxy_tables.rs

local __w1 = "1 true"
local __i = 0

local function ReadOnly(t)
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
do local __t = tostring(ro_data.a .. ' ' .. tostring(string.find(err, 'read-only') ~= nil)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
