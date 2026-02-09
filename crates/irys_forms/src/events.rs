use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, PartialEq, Eq, Hash, Serialize, Deserialize)]
pub enum EventType {
    Click,
    DblClick,
    Load,
    Unload,
    Change,
    KeyPress,
    KeyDown,
    KeyUp,
    MouseDown,
    MouseUp,
    MouseMove,
    GotFocus,
    LostFocus,
}

impl EventType {
    pub fn as_str(&self) -> &str {
        match self {
            EventType::Click => "Click",
            EventType::DblClick => "DblClick",
            EventType::Load => "Load",
            EventType::Unload => "Unload",
            EventType::Change => "Change",
            EventType::KeyPress => "KeyPress",
            EventType::KeyDown => "KeyDown",
            EventType::KeyUp => "KeyUp",
            EventType::MouseDown => "MouseDown",
            EventType::MouseUp => "MouseUp",
            EventType::MouseMove => "MouseMove",
            EventType::GotFocus => "GotFocus",
            EventType::LostFocus => "LostFocus",
        }
    }

    pub fn parameters(&self) -> &'static str {
        match self {
            EventType::KeyPress => "KeyAscii As Integer",
            EventType::KeyDown | EventType::KeyUp => "KeyCode As Integer, Shift As Integer",
            EventType::MouseDown | EventType::MouseUp => "Button As Integer, Shift As Integer, X As Single, Y As Single",
            EventType::MouseMove => "Button As Integer, Shift As Integer, X As Single, Y As Single",
            _ => "",
        }
    }

    pub fn is_applicable_to(&self, control_type: Option<crate::ControlType>) -> bool {
        use crate::ControlType;
        
        match self {
            // Form-only events
            EventType::Load | EventType::Unload => control_type.is_none(),
            
            // Change is for TextBox, ComboBox, ListBox, etc.
            EventType::Change => matches!(
                control_type,
                Some(ControlType::TextBox)
                    | Some(ControlType::Label)
                    | Some(ControlType::ComboBox)
                    | Some(ControlType::ListBox)
            ),
            
            // Most controls support these
            EventType::Click | EventType::DblClick => true,
            
            // Keyboard events - most controls
            EventType::KeyPress | EventType::KeyDown | EventType::KeyUp => true,
            
            // Mouse events - all controls
            EventType::MouseDown | EventType::MouseUp | EventType::MouseMove => true,
            
            // Focus events - most controls
            EventType::GotFocus | EventType::LostFocus => control_type.is_some(),
        }
    }

    pub fn all_events() -> Vec<EventType> {
        vec![
            EventType::Click,
            EventType::DblClick,
            EventType::Load,
            EventType::Unload,
            EventType::Change,
            EventType::KeyPress,
            EventType::KeyDown,
            EventType::KeyUp,
            EventType::MouseDown,
            EventType::MouseUp,
            EventType::MouseMove,
            EventType::GotFocus,
            EventType::LostFocus,
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
