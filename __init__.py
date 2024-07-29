import gradio as gr

from plugins.webui.sanp_plugin_image_compression.utils import (
    compression_and_organization,
    image_compression,
    image_organization,
)


def plugin():
    with gr.Tab("压缩整理"):
        image_path = gr.Textbox("", label="图片目录")
        image_format = gr.Radio(
            ["jpg", "png"],
            value="png",
            label="压缩格式(jpg 格式图片较小, 但无法存储 pnginfo; png 格式反之)",
        )
        with gr.Row():
            compression_button = gr.Button("仅压缩")
            organization_button = gr.Button("仅整理")
            compression_and_organization_button = gr.Button("压缩并整理")
        output_info = gr.Textbox(label="输出信息")

        compression_button.click(
            image_compression, inputs=[image_format, image_path], outputs=output_info
        )
        organization_button.click(
            image_organization, inputs=image_path, outputs=output_info
        )
        compression_and_organization_button.click(
            compression_and_organization,
            inputs=[image_format, image_path],
            outputs=output_info,
        )
