-- vybe-test: lua/metatables_proxy_tables/test_proxy_strict_mode
-- origin: languages/lua/tests/lua/test_metatables_proxy_tables.rs

local __w1 = "true true"
local __i = 0

local function Strict(t)
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
do local __t = tostring(tostring(string.find(err_read, 'undeclared') ~= nil) .. ' ' .. tostring(string.find(err_write, 'undeclared') ~= nil)); __i = __i + 1
  if __i == 1 and __t ~= __w1 then error("FAIL: want [" .. __w1 .. "] got [" .. __t .. "]") end end

if __i == 0 then error("FAIL: no output, wanted [" .. __w1 .. "]") end
