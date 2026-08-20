-- App 09: RPG Combat Simulator (prototype inheritance)

local Entity = {}
Entity.__index = Entity

function Entity.new(name, hp, atk)
  return setmetatable({name = name, hp = hp, atk = atk, effects = {}}, Entity)
end

function Entity:is_alive() return self.hp > 0 end
function Entity:take_damage(dmg) self.hp = math.max(0, self.hp - dmg) end
function Entity:add_effect(name, turns) self.effects[name] = turns end
function Entity:tick_effects()
  for k, v in pairs(self.effects) do
    self.effects[k] = v - 1
    if self.effects[k] <= 0 then self.effects[k] = nil end
  end
end

local Mage = setmetatable({}, {__index = Entity})
Mage.__index = Mage

function Mage.new(name)
  local self = Entity.new(name, 85, 14)
  self.mana = 100
  return setmetatable(self, Mage)
end

function Mage:cast_fireball(target)
  if self.mana < 25 then return false end
  self.mana = self.mana - 25
  target:take_damage(self.atk + 12)
  target:add_effect("burn", 2)
  return true
end

local hero = Mage.new("Lyra")
local boss = Entity.new("Golem", 160, 10)

for turn = 1, 8 do
  if not (hero:is_alive() and boss:is_alive()) then break end
  if not hero:cast_fireball(boss) then boss:take_damage(hero.atk) end
  if boss.effects.burn then boss:take_damage(4) end
  boss:tick_effects()
  if boss:is_alive() then hero:take_damage(boss.atk) end
  print(string.format("turn %d -> hero hp=%d mana=%d | boss hp=%d", turn, hero.hp, hero.mana, boss.hp))
end

print("winner:", hero:is_alive() and "hero" or "boss")
