-- 12_file_io_and_serialization.lua
-- Demonstrates file I/O and simple table serialization.

local function serialize(tbl, indent)
  indent = indent or ""
  local lines = {"{"}
  local next_indent = indent .. "  "

  for k, v in pairs(tbl) do
    local key = string.format("[%q]", tostring(k))
    local value
    if type(v) == "table" then
      value = serialize(v, next_indent)
    elseif type(v) == "string" then
      value = string.format("%q", v)
    else
      value = tostring(v)
    end
    lines[#lines + 1] = next_indent .. key .. " = " .. value .. ","
  end

  lines[#lines + 1] = indent .. "}"
  return table.concat(lines, "\n")
end

local profile = {
  name = "Nora",
  level = 7,
  flags = {hardcore = false, mentor = true}
}

local filename = "lua_profile_dump.txt"
local f = assert(io.open(filename, "w"))
f:write(serialize(profile), "\n")
f:close()

local rf = assert(io.open(filename, "r"))
local content = rf:read("*a")
rf:close()

print("written and read:")
print(content)

local ok, remove_err = os.remove(filename)
print("cleanup:", ok, remove_err)
