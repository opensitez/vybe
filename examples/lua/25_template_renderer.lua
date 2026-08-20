-- App 25: Template Renderer with Filters

local filters = {
  upper = function(s) return string.upper(s) end,
  lower = function(s) return string.lower(s) end,
  money = function(n) return string.format("$%.2f", tonumber(n) or 0) end,
}

local function render(template, ctx)
  return (template:gsub("{{%s*([%w_]+)%s*|?%s*([%w_]*)%s*}}", function(key, filter)
    local value = ctx[key]
    if value == nil then return "" end
    if filter ~= "" and filters[filter] then
      value = filters[filter](value)
    end
    return tostring(value)
  end))
end

local invoice = {
  customer = "Acme Corp",
  item = "Support Plan",
  total = 1299.5,
  due = "2026-07-01",
}

local tpl = [[
Invoice for {{ customer | upper }}
Item: {{ item }}
Total: {{ total | money }}
Due: {{ due }}
]]

print(render(tpl, invoice))
