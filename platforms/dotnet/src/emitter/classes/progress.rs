//! Progress / value-input controls.
//!
//! `ProgressBar`, `TrackBar`, and `NumericUpDown` all inherit from
//! `Control` directly in real .NET (`NumericUpDown` actually inherits
//! from `UpDownBase` which inherits from `ContainerControl`, but for our
//! property-binding purposes we keep it shallow until/if a user needs
//! `UpDownBase`-specific subclassing).

use super::DotnetClass;

pub fn classes() -> &'static [DotnetClass] {
    &[
        DotnetClass {
            name: "ProgressBar",
            parent: Some("Control"),
            properties: &[
                "MarqueeAnimationSpeed",
                "Maximum",
                "Minimum",
                "Step",
                "Style",
                "Value",
            ],
            methods: &[],
            ctor_arity: 0,
            // `<progress>` — created by the element mapping.
            widget_host_fn: None,
            widget_host_module: "vybe:gui",
        },
        DotnetClass {
            name: "TrackBar",
            parent: Some("Control"),
            properties: &[
                "AutoSize",
                "LargeChange",
                "Maximum",
                "Minimum",
                "Orientation",
                "SmallChange",
                "TickFrequency",
                "TickStyle",
                "Value",
            ],
            methods: &[],
            ctor_arity: 0,
            // `<input type="range">`
            widget_host_fn: None,
            widget_host_module: "vybe:gui",
        },
        DotnetClass {
            name: "NumericUpDown",
            parent: Some("Control"),
            properties: &[
                "AutoSize",
                "DecimalPlaces",
                "Hexadecimal",
                "Increment",
                "Maximum",
                "Minimum",
                "ReadOnly",
                "ThousandsSeparator",
                "UpDownAlign",
                "Value",
            ],
            methods: &[],
            ctor_arity: 0,
            // `<input type="number">`
            widget_host_fn: None,
            widget_host_module: "vybe:gui",
        },
    ]
}
