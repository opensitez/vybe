-- App 02: Pricing Engine (coupon strategy pipeline)

local cart = {
  {sku = "BOOK", qty = 2, unit = 15.00, category = "edu"},
  {sku = "PEN", qty = 5, unit = 2.50, category = "office"},
  {sku = "LAPTOP", qty = 1, unit = 999.00, category = "tech"},
}

local function subtotal(lines)
  local sum = 0
  for _, l in ipairs(lines) do sum = sum + l.qty * l.unit end
  return sum
end

local function compose(...)
  local fns = {...}
  return function(state)
    for _, fn in ipairs(fns) do state = fn(state) end
    return state
  end
end

local function category_discount(cat, pct)
  return function(state)
    local saved = 0
    for _, l in ipairs(state.lines) do
      if l.category == cat then
        saved = saved + l.qty * l.unit * pct
      end
    end
    state.discounts[#state.discounts + 1] = {name = cat .. " promo", value = saved}
    state.total = state.total - saved
    return state
  end
end

local function threshold_discount(min_total, pct)
  return function(state)
    if state.total >= min_total then
      local saved = state.total * pct
      state.discounts[#state.discounts + 1] = {name = "threshold", value = saved}
      state.total = state.total - saved
    end
    return state
  end
end

local checkout = compose(
  category_discount("office", 0.10),
  threshold_discount(200, 0.05)
)

local state = {lines = cart, total = subtotal(cart), discounts = {}}
state = checkout(state)

print("Final total:", string.format("%.2f", state.total))
for _, d in ipairs(state.discounts) do
  print("discount:", d.name, string.format("%.2f", d.value))
end
