-- App 10: Plugin-driven Command Console (module registry)

package.preload["plugins.math_ops"] = function()
  return {
    name = "math",
    commands = {
      add = function(args) return tonumber(args[1]) + tonumber(args[2]) end,
      mul = function(args) return tonumber(args[1]) * tonumber(args[2]) end,
    }
  }
end

package.preload["plugins.text_ops"] = function()
  return {
    name = "text",
    commands = {
      upper = function(args) return string.upper(table.concat(args, " ")) end,
      slug = function(args)
        local s = table.concat(args, " "):lower():gsub("%s+", "-"):gsub("[^%w%-]", "")
        return s
      end,
    }
  }
end

local plugins = {
  require("plugins.math_ops"),
  require("plugins.text_ops"),
}

local registry = {}
for _, p in ipairs(plugins) do
  for cmd, fn in pairs(p.commands) do
    registry[cmd] = fn
  end
end

local script = {
  "add 20 22",
  "mul 7 6",
  "upper lua plugin systems are neat",
  "slug Hello Lua, Plugin World!",
}

for _, line in ipairs(script) do
  local parts = {}
  for token in line:gmatch("%S+") do parts[#parts + 1] = token end
  local cmd = table.remove(parts, 1)
  local fn = registry[cmd]
  if fn then
    print(cmd, "=>", fn(parts))
  else
    print("unknown command:", cmd)
  end
end
