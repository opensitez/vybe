//! random module: randint, randrange, uniform, seed, choices, sample edge cases.

crate::runtime_case!(
    random_choice_from_list,
    "import random\nprint(random.choice([7]))\n",
    "7"
);
crate::runtime_case!(
    random_choice_from_tuple,
    "import random\nprint(random.choice((10, 20)) in (10, 20))\n",
    "True"
);
crate::runtime_case!(
    random_choice_from_string,
    "import random\nprint(random.choice('ab') in 'ab')\n",
    "True"
);
crate::runtime_case!(
    random_shuffle_preserves_len,
    "import random\na = [1, 2, 3, 4, 5]\nrandom.shuffle(a)\nprint(len(a))\n",
    "5"
);
crate::runtime_case!(
    random_shuffle_same_elements,
    "import random\na = [1, 2, 3]\nrandom.shuffle(a)\nprint(sorted(a))\n",
    "[1, 2, 3]"
);
crate::runtime_case!(
    random_sample_size,
    "import random\nr = random.sample([1, 2, 3, 4, 5], 3)\nprint(len(r))\n",
    "3"
);
crate::runtime_case!(
    random_sample_no_duplicates,
    "import random\nr = random.sample([1, 2, 3], 3)\nprint(len(set(r)))\n",
    "3"
);
crate::runtime_case!(
    random_randint_in_range,
    "import random\nn = random.randint(5, 5)\nprint(n)\n",
    "5"
);
crate::runtime_case!(
    random_randrange_stop,
    "import random\nn = random.randrange(3)\nprint(0 <= n < 3)\n",
    "True"
);
crate::runtime_case!(
    random_randrange_start_stop,
    "import random\nn = random.randrange(10, 12)\nprint(n in (10, 11))\n",
    "True"
);
crate::runtime_case!(
    random_randrange_step,
    "import random\nn = random.randrange(0, 10, 2)\nprint(n % 2 == 0)\n",
    "True"
);
crate::runtime_case!(
    random_uniform_is_float,
    "import random\nx = random.uniform(1.0, 2.0)\nprint(isinstance(x, float))\n",
    "True"
);
crate::runtime_case!(
    random_random_in_unit_interval,
    "import random\nx = random.random()\nprint(0.0 <= x < 1.0)\n",
    "True"
);
crate::runtime_case!(
    random_seed_reproducible,
    "import random\nrandom.seed(42)\na = random.randint(1, 100)\nrandom.seed(42)\nb = random.randint(1, 100)\nprint(a == b)\n",
    "True"
);
crate::runtime_case!(
    random_getstate_setstate,
    "import random\nrandom.seed(1)\nst = random.getstate()\nr1 = random.random()\nrandom.setstate(st)\nr2 = random.random()\nprint(r1 == r2)\n",
    "True"
);
crate::runtime_case!(
    random_getrandbits,
    "import random\nn = random.getrandbits(4)\nprint(0 <= n < 16)\n",
    "True"
);
crate::runtime_case!(
    random_choices_with_weights,
    "import random\nr = random.choices(['a', 'b'], weights=[1, 0], k=3)\nprint(r == ['a', 'a', 'a'])\n",
    "True"
);
crate::runtime_case!(
    random_choices_k_len,
    "import random\nr = random.choices([1, 2, 3], k=5)\nprint(len(r))\n",
    "5"
);
crate::runtime_case!(
    random_sample_population_range,
    "import random\nr = random.sample(range(100), 5)\nprint(len(r))\n",
    "5"
);
crate::runtime_case!(
    random_triangular_default,
    "import random\nx = random.triangular(0, 10)\nprint(0 <= x <= 10)\n",
    "True"
);
crate::runtime_case!(
    random_betavariate,
    "import random\nx = random.betavariate(2, 5)\nprint(0 <= x <= 1)\n",
    "True"
);
crate::runtime_case!(
    random_expovariate,
    "import random\nx = random.expovariate(1.0)\nprint(x >= 0)\n",
    "True"
);
crate::runtime_case!(
    random_gammavariate,
    "import random\nx = random.gammavariate(2, 1)\nprint(x >= 0)\n",
    "True"
);
crate::runtime_case!(
    random_gauss,
    "import random\nx = random.gauss(0, 1)\nprint(isinstance(x, float))\n",
    "True"
);
crate::runtime_case!(
    random_lognormvariate,
    "import random\nx = random.lognormvariate(0, 1)\nprint(x > 0)\n",
    "True"
);
crate::runtime_case!(
    random_normalvariate,
    "import random\nx = random.normalvariate(0, 1)\nprint(isinstance(x, float))\n",
    "True"
);
crate::runtime_case!(
    random_vonmisesvariate,
    "import random\nx = random.vonmisesvariate(0, 1)\nprint(isinstance(x, float))\n",
    "True"
);
crate::runtime_case!(
    random_paretovariate,
    "import random\nx = random.paretovariate(3)\nprint(x >= 1)\n",
    "True"
);
crate::runtime_case!(
    random_weibullvariate,
    "import random\nx = random.weibullvariate(1, 1)\nprint(x >= 0)\n",
    "True"
);
crate::runtime_case!(
    random_seed_none,
    "import random\nrandom.seed()\nx = random.random()\nprint(isinstance(x, float))\n",
    "True"
);
crate::runtime_case!(
    random_seed_int,
    "import random\nrandom.seed(0)\nprint(isinstance(random.random(), float))\n",
    "True"
);
crate::runtime_case!(
    random_seed_str,
    "import random\nrandom.seed('hello')\nprint(isinstance(random.random(), float))\n",
    "True"
);
crate::runtime_case!(
    random_sample_set_population,
    "import random\nr = random.sample({1, 2, 3, 4}, 2)\nprint(len(r))\n",
    "2"
);
crate::runtime_case!(
    random_shuffle_inplace,
    "import random\na = [3, 1, 2]\nrandom.shuffle(a)\nprint(len(a) == 3)\n",
    "True"
);
crate::runtime_case!(
    random_choice_sequence_protocol,
    "import random\nprint(random.choice(bytearray(b'xy')) in (120, 121))\n",
    "True"
);
crate::runtime_case!(
    random_randbytes,
    "import random\nb = random.randbytes(4)\nprint(len(b))\n",
    "4"
);
crate::runtime_case!(
    random_sample_count_one,
    "import random\nprint(random.sample([9], 1))\n",
    "[9]"
);
crate::runtime_case!(
    random_randrange_negative_step,
    "import random\nn = random.randrange(10, 0, -2)\nprint(n % 2 == 0)\n",
    "True"
);
crate::runtime_case!(
    random_random_after_seed_zero,
    "import random\nrandom.seed(0)\nrandom.random()\nrandom.random()\nprint(isinstance(random.random(), float))\n",
    "True"
);
crate::runtime_case!(
    random_choices_empty_weight_sum,
    "import random\nr = random.choices('ab', k=1)\nprint(len(r))\n",
    "1"
);
crate::runtime_case!(
    random_uniform_same_bounds,
    "import random\nprint(random.uniform(3, 3))\n",
    "3.0"
);
crate::runtime_case!(
    random_getstate_is_tuple,
    "import random\nprint(isinstance(random.getstate(), tuple))\n",
    "True"
);
crate::runtime_case!(
    random_module_has_random_func,
    "import random\nprint(callable(random.random))\n",
    "True"
);
crate::runtime_case!(
    random_module_has_seed,
    "import random\nprint(callable(random.seed))\n",
    "True"
);
crate::runtime_case!(
    random_randint_swapped_raises,
    "import random\ntry:\n    random.randint(10, 5)\n    print('ok')\nexcept ValueError:\n    print('bad')\n",
    "bad"
);

crate::compile_case!(
    random_sample_k_equals_n,
    "import random\nrandom.sample([1, 2, 3], 3)\n"
);
crate::compile_case!(
    random_choices_cum_weights,
    "import random\nrandom.choices([1, 2], cum_weights=[1, 3], k=2)\n"
);
crate::compile_case!(
    random_shuffle_copy,
    "import random\na = [1, 2]\nrandom.shuffle(a)\n"
);
crate::compile_case!(
    random_randrange_zero_step_raises,
    "import random\ntry:\n    random.randrange(0, 10, 0)\nexcept ValueError:\n    pass\n"
);
crate::compile_case!(
    random_sample_k_too_large,
    "import random\ntry:\n    random.sample([1, 2], 3)\nexcept ValueError:\n    pass\n"
);
