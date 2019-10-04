import dash
import dash_table
import dash_core_components as dcc
import dash_html_components as html
from dash.dependencies import Input, Output, State
from powerpoint_analyzer import PowerpointAnalyzer
import base64
import datetime
import io


powerpoint_analyzer = PowerpointAnalyzer()

app = dash.Dash(__name__)

app.layout = html.Div([
    html.H2("Powerpoint Analyzer"),
    dcc.Upload(
        id='upload-data',
        children=html.Div([
            'Drag and Drop or ',
            html.A('Select Files')
        ]),
        style={
            'width': '100%',
            'height': '60px',
            'lineHeight': '60px',
            'borderWidth': '1px',
            'borderStyle': 'dashed',
            'borderRadius': '5px',
            'textAlign': 'center',
            'margin': '10px'
        },
        # Allow multiple files to be uploaded
        multiple=True
    ),
    html.Div(id='output-data-upload'),
])


def parse_contents(content, filename, date):
    data = content.encode("utf8").split(b";base64,")[1]
    decoded = base64.decodebytes(data)
    f = io.BytesIO(decoded)

    # try:
    if '.ppt' in filename:
        # Assume that the user uploaded a Powerpoint
        fonts_data = powerpoint_analyzer.analyze_fonts(f)
    else:
        return html.Div(
            [
                'File is not a Powerpoint'
            ],
            className="error"
        )
    # except Exception as e:
    #     print(str(e))
    #     return html.Div([
    #         'There was an error processing this file.'
    #     ])

    return html.Div([
        html.H5(filename),
        html.H6(datetime.datetime.fromtimestamp(date)),

        html.P(f"Fonts : {list(fonts_data.keys())}"),

        *[
            html.Div(
                [
                    html.H5(str(font)),
                    html.P(f"Count : {font_data['count']}"),
                    html.P("Used on slides : "),
                    html.Ul(
                        [
                            html.Li(slide_idx)
                            for slide_idx in font_data["slides"]
                        ]
                    )
                ]
            )
            for font, font_data in fonts_data.items()
        ]

    ])


@app.callback(
    Output('output-data-upload', 'children'),
    [Input('upload-data', 'contents')],
    [
        State('upload-data', 'filename'),
        State('upload-data', 'last_modified')
    ]
)
def update_output(content, name, date):
    if content is not None:
        children = [
            parse_contents(content[0], name[0], date[0])
        ]
        return children


if __name__ == '__main__':
    app.run_server(
        debug=True,
        host='0.0.0.0'
    )
