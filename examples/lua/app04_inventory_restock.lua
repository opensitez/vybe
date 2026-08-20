-- Full App 04: Inventory Restock Simulator with reorder policies

local sku = {
  {id = "WID-A", stock = 35, daily_mean = 8, lead = 3, reorder_point = 28, reorder_qty = 60},
  {id = "WID-B", stock = 16, daily_mean = 5, lead = 4, reorder_point = 20, reorder_qty = 40},
  {id = "WID-C", stock = 80, daily_mean = 12, lead = 2, reorder_point = 35, reorder_qty = 90},
}

local inbound = {}

local function receive(day)
  for i = #inbound, 1, -1 do
    local s = inbound[i]
    if s.day == day then
      s.item.stock = s.item.stock + s.qty
      print(string.format("day %d receive %s +%d stock=%d", day, s.item.id, s.qty, s.item.stock))
      table.remove(inbound, i)
    end
  end
end

local function consume(item)
  local demand = math.max(0, item.daily_mean + ((item.stock % 3) - 1))
  item.stock = math.max(0, item.stock - demand)
end

local function maybe_reorder(day, item)
  local has_open = false
  for _, s in ipairs(inbound) do
    if s.item == item then has_open = true end
  end
  if not has_open and item.stock <= item.reorder_point then
    inbound[#inbound + 1] = {day = day + item.lead, item = item, qty = item.reorder_qty}
    print(string.format("day %d reorder %s qty=%d eta=%d", day, item.id, item.reorder_qty, day + item.lead))
  end
end

for day = 1, 14 do
  receive(day)
  for _, item in ipairs(sku) do
    consume(item)
    maybe_reorder(day, item)
  end
end

for _, item in ipairs(sku) do
  print(item.id, "final stock=", item.stock)
end
