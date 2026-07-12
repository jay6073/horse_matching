import reflex as rx

config = rx.Config(
    app_name="horse_matching",
    plugins=[
        rx.plugins.SitemapPlugin(),
        rx.plugins.TailwindV4Plugin(),
    ]
)