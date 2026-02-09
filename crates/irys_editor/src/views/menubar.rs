use crate::message::Message;
use iced::widget::{button, row, text};
use iced::Element;

pub fn menubar() -> Element<'static, Message> {
    row![
        button(text("New Project")).on_press(Message::NewProject),
        button(text("New Form")).on_press(Message::NewForm),
        button(text("Save")).on_press(Message::SaveProject),
        button(text("Generate Code")).on_press(Message::GenerateEventHandlers),
        button(text("Run")).on_press(Message::Start),
    ]
    .spacing(5)
    .padding(5)
    .into()
}
