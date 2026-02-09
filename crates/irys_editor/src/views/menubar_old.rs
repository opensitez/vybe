use crate::message::Message;
use iced::widget::{button, container, row, text};
use iced::{Element, Length};

pub fn menubar() -> Element<'static, Message> {
    let menu = row![
        button(text("📄 New").size(13))
            .on_press(Message::NewProject)
            .padding(8)
            .style(menu_button_style),
        button(text("📁 Open").size(13))
            .on_press(Message::OpenProject)
            .padding(8)
            .style(menu_button_style),
        button(text("💾 Save").size(13))
            .on_press(Message::SaveProject)
            .padding(8)
            .style(menu_button_style),
        container(text(""))
            .width(Length::Fixed(10.0)),
        button(text("▶ Run").size(13))
            .on_press(Message::RunProject)
            .padding(8)
            .style(run_button_style),
    ]
    .spacing(5)
    .padding(8);

    container(menu)
        .width(Length::Fill)
        .style(|_theme| {
            container::Style {
                background: Some(iced::Background::Color(iced::Color::from_rgb(0.94, 0.94, 0.94))),
                border: iced::Border {
                    color: iced::Color::from_rgb(0.8, 0.8, 0.8),
                    width: 1.0,
                    ..Default::default()
                },
                ..Default::default()
            }
        })
        .into()
}

fn menu_button_style(_theme: &iced::Theme, status: button::Status) -> button::Style {
    button::Style {
        background: Some(iced::Background::Color(
            match status {
                button::Status::Hovered => iced::Color::from_rgb(0.85, 0.85, 0.85),
                button::Status::Pressed => iced::Color::from_rgb(0.75, 0.75, 0.75),
                _ => iced::Color::from_rgb(0.94, 0.94, 0.94),
            }
        )),
        border: iced::Border {
            color: match status {
                button::Status::Hovered | button::Status::Pressed => iced::Color::from_rgb(0.6, 0.6, 0.6),
                _ => iced::Color::from_rgb(0.8, 0.8, 0.8),
            },
            width: 1.0,
            radius: 3.0.into(),
        },
        text_color: iced::Color::BLACK,
        ..Default::default()
    }
}

fn run_button_style(_theme: &iced::Theme, status: button::Status) -> button::Style {
    button::Style {
        background: Some(iced::Background::Color(
            match status {
                button::Status::Hovered => iced::Color::from_rgb(0.2, 0.7, 0.3),
                button::Status::Pressed => iced::Color::from_rgb(0.15, 0.6, 0.25),
                _ => iced::Color::from_rgb(0.25, 0.75, 0.35),
            }
        )),
        border: iced::Border {
            color: iced::Color::from_rgb(0.2, 0.6, 0.3),
            width: 1.0,
            radius: 3.0.into(),
        },
        text_color: iced::Color::WHITE,
        ..Default::default()
    }
}
