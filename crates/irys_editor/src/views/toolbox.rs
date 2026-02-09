use crate::message::Message;
use iced::widget::{button, column, container, text};
use iced::{Element, Length};
use irys_forms::ControlType;

pub fn toolbox() -> Element<'static, Message> {
    let controls = vec![
        ControlType::Button,
        ControlType::Label,
        ControlType::TextBox,
        ControlType::CheckBox,
        ControlType::RadioButton,
        ControlType::RichTextBox,
        ControlType::WebBrowser,
    ];

    let mut col = column![text("Toolbox").size(16)]
        .spacing(5)
        .padding(10);

    for control_type in controls {
        col = col.push(
            button(text(control_type.as_str()))
                .width(Length::Fill)
                .on_press(Message::SelectTool(control_type))
        );
    }

    container(col)
        .width(Length::Fixed(150.0))
        .height(Length::Fill)
        .into()
}
