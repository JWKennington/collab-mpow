'''
Some useful plotting functions
'''


from bokeh.io import output_notebook, show
from bokeh.plotting import figure
import numpy
import pandas


def histogram(data: pandas.Series, bins: int=10, title: str=None, density: bool=False, width: int=600, height: int=300):
    counts, edges = numpy.histogram(data.values, bins=bins)
    centers = (edges[:-1] + edges[1:]) / 2
    if density:
        counts = counts / numpy.sum(counts)

    fig = figure(title='Histogram: {}'.format(data.name) if title is None else title, plot_width=width, plot_height=height)
    fig.line(x=centers, y=counts)
    return fig