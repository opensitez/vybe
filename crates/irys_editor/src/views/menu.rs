use crate::message::Message;
use iced::widget::{button, column, container, row, text};
use iced::{Element, Length};

pub fn menu_bar() -> Element<'static, Message> {
    row![
        file_menu(),
        edit_menu(),
        view_menu(),
        project_menu(),
        run_menu(),
        window_menu(),
    ]
    .spacing(2)
    .padding(2)
    .into()
}

fn file_menu() -> Element<'static, Message> {
    // In a full implementation, this would be a dropdown
    // For now, showing key items
    button(text("File").size(13))
        .on_press(Message::None)
        .into()
}

fn edit_menu() -> Element<'static, Message> {
    button(text("Edit").size(13))
        .on_press(Message::None)
        .into()
}

fn view_menu() -> Element<'static, Message> {
    button(text("View").size(13))
        .on_press(Message::None)
        .into()
}

fn project_menu() -> Element<'static, Message> {
    button(text("Project").size(13))
        .on_press(Message::None)
        .into()
}

fn run_menu() -> Element<'static, Message> {
    button(text("Run").size(13))
        .on_press(Message::None)
        .into()
}

fn window_menu() -> Element<'static, Message> {
    button(text("Window").size(13))
        .on_press(Message::None)
        .into()
}

// Toolbar with common actions
pub fn toolbar() -> Element<'static, Message> {
    row![
        button(text("📄").size(16)).on_press(Message::NewProject),
        button(text("📁").size(16)).on_press(Message::OpenProject),
        button(text("💾").size(16)).on_press(Message::SaveProject),
        text("|").size(16),
        button(text("✂️").size(16)).on_press(Message::Cut),
        button(text("📋").size(16)).on_press(Message::Copy),
        button(text("📄").size(16)).on_press(Message::Paste),
        text("|").size(16),
        button(text("▶️").size(16)).on_press(Message::Start),
        button(text("⏸️").size(16)).on_press(Message::Stop),
        text("|").size(16),
        button(text("➕ Form").size(13)).on_press(Message::AddForm),
    ]
    .spacing(5)
    .padding(5)
    .into()
}
