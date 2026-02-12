use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum EventType {
    Click,
    DblClick,
    DoubleClick,
    Load,
    Unload,
    Change,
    TextChanged,
    SelectedIndexChanged,
    CheckedChanged,
    ValueChanged,
    KeyPress,
    KeyDown,
    KeyUp,
    MouseClick,
    MouseDoubleClick,
    MouseDown,
    MouseUp,
    MouseMove,
    MouseEnter,
    MouseLeave,
    MouseWheel,
    GotFocus,
    LostFocus,
    Enter,
    Leave,
    Validated,
    Validating,
    Resize,
    Paint,
    FormClosing,
    FormClosed,
    Shown,
    Activated,
    Deactivate,
    Tick,
    Elapsed,
    Scroll,
    SelectedValueChanged,
    CellClick,
    CellDoubleClick,
    CellValueChanged,
    SelectionChanged,
}

impl EventType {
    pub fn as_str(&self) -> &str {
        match self {
            EventType::Click => "Click",
            EventType::DblClick => "DblClick",
            EventType::DoubleClick => "DoubleClick",
            EventType::Load => "Load",
            EventType::Unload => "Unload",
            EventType::Change => "Change",
            EventType::TextChanged => "TextChanged",
            EventType::SelectedIndexChanged => "SelectedIndexChanged",
            EventType::CheckedChanged => "CheckedChanged",
            EventType::ValueChanged => "ValueChanged",
            EventType::KeyPress => "KeyPress",
            EventType::KeyDown => "KeyDown",
            EventType::KeyUp => "KeyUp",
            EventType::MouseClick => "MouseClick",
            EventType::MouseDoubleClick => "MouseDoubleClick",
            EventType::MouseDown => "MouseDown",
            EventType::MouseUp => "MouseUp",
            EventType::MouseMove => "MouseMove",
            EventType::MouseEnter => "MouseEnter",
            EventType::MouseLeave => "MouseLeave",
            EventType::MouseWheel => "MouseWheel",
            EventType::GotFocus => "GotFocus",
            EventType::LostFocus => "LostFocus",
            EventType::Enter => "Enter",
            EventType::Leave => "Leave",
            EventType::Validated => "Validated",
            EventType::Validating => "Validating",
            EventType::Resize => "Resize",
            EventType::Paint => "Paint",
            EventType::FormClosing => "FormClosing",
            EventType::FormClosed => "FormClosed",
            EventType::Shown => "Shown",
            EventType::Activated => "Activated",
            EventType::Deactivate => "Deactivate",
            EventType::Tick => "Tick",
            EventType::Elapsed => "Elapsed",
            EventType::Scroll => "Scroll",
            EventType::SelectedValueChanged => "SelectedValueChanged",
            EventType::CellClick => "CellClick",
            EventType::CellDoubleClick => "CellDoubleClick",
            EventType::CellValueChanged => "CellValueChanged",
            EventType::SelectionChanged => "SelectionChanged",
        }
    }

    /// Return the .NET-compatible parameter signature for event handlers.
    pub fn parameters(&self) -> &'static str {
        match self {
            // Mouse events use MouseEventArgs
            EventType::MouseClick | EventType::MouseDoubleClick |
            EventType::MouseDown | EventType::MouseUp | EventType::MouseMove |
            EventType::MouseWheel =>
                "sender As Object, e As MouseEventArgs",
            // Key events
            EventType::KeyDown | EventType::KeyUp =>
                "sender As Object, e As KeyEventArgs",
            EventType::KeyPress =>
                "sender As Object, e As KeyPressEventArgs",
            // Form closing
            EventType::FormClosing =>
                "sender As Object, e As FormClosingEventArgs",
            EventType::FormClosed =>
                "sender As Object, e As FormClosedEventArgs",
            // Paint
            EventType::Paint =>
                "sender As Object, e As PaintEventArgs",
            // All other events use base EventArgs
            _ => "sender As Object, e As EventArgs",
        }
    }

    pub fn is_applicable_to(&self, control_type: Option<crate::ControlType>) -> bool {
        use crate::ControlType;
        
        match self {
            // Form-only events
            EventType::Load | EventType::Unload | EventType::FormClosing | EventType::FormClosed
            | EventType::Shown | EventType::Activated | EventType::Deactivate => control_type.is_none(),
            
            // Text change events: TextBox, ComboBox, ListBox, etc.
            EventType::Change | EventType::TextChanged => matches!(
                control_type,
                Some(ControlType::TextBox)
                    | Some(ControlType::Label)
                    | Some(ControlType::ComboBox)
                    | Some(ControlType::ListBox)
            ),
            
            // Selection events
            EventType::SelectedIndexChanged | EventType::SelectedValueChanged => matches!(
                control_type,
                Some(ControlType::ComboBox) | Some(ControlType::ListBox)
            ),

            // Checked state
            EventType::CheckedChanged => matches!(
                control_type,
                Some(ControlType::CheckBox) | Some(ControlType::RadioButton)
            ),

            // Value changed
            EventType::ValueChanged => true,
            
            // Timer events - handled at runtime level, not via visual controls
            EventType::Tick | EventType::Elapsed => true,

            // DataGridView events
            EventType::CellClick | EventType::CellDoubleClick | EventType::CellValueChanged
            | EventType::SelectionChanged => matches!(
                control_type,
                Some(ControlType::DataGridView)
            ),
            
            // Most controls support click/double-click
            EventType::Click | EventType::DblClick | EventType::DoubleClick | EventType::MouseClick | EventType::MouseDoubleClick => true,
            
            // Keyboard events
            EventType::KeyPress | EventType::KeyDown | EventType::KeyUp => true,
            
            // Mouse events - all controls
            EventType::MouseDown | EventType::MouseUp | EventType::MouseMove
            | EventType::MouseEnter | EventType::MouseLeave | EventType::MouseWheel => true,
            
            // Focus events
            EventType::GotFocus | EventType::LostFocus | EventType::Enter | EventType::Leave
            | EventType::Validated | EventType::Validating => control_type.is_some(),

            // Layout/paint
            EventType::Resize | EventType::Paint | EventType::Scroll => true,
        }
    }

    pub fn all_events() -> Vec<EventType> {
        vec![
            EventType::Click,
            EventType::DblClick,
            EventType::DoubleClick,
            EventType::Load,
            EventType::Unload,
            EventType::Change,
            EventType::TextChanged,
            EventType::SelectedIndexChanged,
            EventType::CheckedChanged,
            EventType::ValueChanged,
            EventType::KeyPress,
            EventType::KeyDown,
            EventType::KeyUp,
            EventType::MouseClick,
            EventType::MouseDoubleClick,
            EventType::MouseDown,
            EventType::MouseUp,
            EventType::MouseMove,
            EventType::MouseEnter,
            EventType::MouseLeave,
            EventType::MouseWheel,
            EventType::GotFocus,
            EventType::LostFocus,
            EventType::Enter,
            EventType::Leave,
            EventType::Validated,
            EventType::Validating,
            EventType::Resize,
            EventType::Paint,
            EventType::FormClosing,
            EventType::FormClosed,
            EventType::Shown,
            EventType::Activated,
            EventType::Deactivate,
            EventType::Tick,
            EventType::Elapsed,
            EventType::Scroll,
            EventType::SelectedValueChanged,
            EventType::CellClick,
            EventType::CellDoubleClick,
            EventType::CellValueChanged,
            EventType::SelectionChanged,
        ]
    }
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct EventBinding {
    pub control_name: String,
    pub event_type: EventType,
    pub handler_name: String,
}

impl EventBinding {
    pub fn new(control_name: impl Into<String>, event_type: EventType) -> Self {
        let control_name = control_name.into();
        let handler_name = format!("{}_{}", control_name, event_type.as_str());

        Self {
            control_name,
            event_type,
            handler_name,
        }
    }

    pub fn with_handler(control_name: impl Into<String>, event_type: EventType, handler_name: impl Into<String>) -> Self {
        Self {
            control_name: control_name.into(),
            event_type,
            handler_name: handler_name.into(),
        }
    }
}
