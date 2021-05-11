import highchartexport as hc_export

bubble_chart_config = {
    "chart": {
        "type": 'bubble',
        "plotBorderWidth": 1,
        "zoomType": 'xy'
    },

    "title": {
        "text": 'Highcharts bubbles save in pdf file'
    },

    "series": [{
        "data": [
            [19, 81, 63],
            [98, 5, 89],
            [51, 50, 73],
        ],
        "marker": {
            "fillColor": {
                "radialGradient": { "cx": 0.4, "cy": 0.3, "r": 0.7 },
                "stops": [
                    [0, 'rgba(255,255,255,0.5)'],
                    [1, 'rgba(200, 200, 200, 0.8)']
                ]
            }
        }
    }]
}

bar_chart_config = {
    "chart": {
        "type": "bar",
        "spacingBottom": 5,
        "spacingTop": 5,
        "spacingLeft": 0,
        "spacingRight": 1
    },
    "tooltip": {
        "valueSuffix": "%"
    },
    "yAxis": {
        "min": 0,
        "title": {
            "text": "Percent of Household"
        }
    },
    "title": {
        "text": "Access To Electricity (Badimalika)"
    },
    "xAxis": {
        "categories": [ "Minigrid", "None", "Solar Home System", "Solar Lantern" ]
    },
    "series": [
        {
            "name": "Household Percent",
            "data": [ 59.2, 26.87, 5.97, 7.96 ],
            "color": "#23527A",
            "showInLegend": False
        }
    ]
}


def generate_all_charts():
    hc_export.save_as_jpeg(config=bubble_chart_config, filename="charts/BUBBLE_CHART.jpeg", width=800)
    hc_export.save_as_jpeg(config=bar_chart_config, filename="charts/EA_CHART.jpeg", width=800, scale=4)


if __name__=="__main__":
    generate_all_charts()
