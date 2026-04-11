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
            widget_host_fn: Some("new_ProgressBar"),
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
            widget_host_fn: Some("new_TrackBar"),
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
            widget_host_fn: Some("new_NumericUpDown"),
        },
    ]
}
