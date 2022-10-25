import os
from openpyxl import Workbook, load_workbook
from scipy.optimize import curve_fit
import numpy as np

directory = os.getcwd()
wb = load_workbook(f'{directory}\\kappa_concentration.xlsx')

ws = wb.active

def linear(x, k):
    return k*x

concentrations = np.array([1/300, 1/400, 1/600, 1/800, 1/1600])

wavelengths = []
kappas = []

for row in ws.iter_rows(min_row=2):
    wavelength = float(str(row[0].value).replace(',', '.'))
    kappa_values = np.array([float(str(raw_cell_data.value).replace(',', '.')) for raw_cell_data in row[1:]])

    pars, _ = curve_fit(f=linear, xdata=concentrations, ydata=kappa_values)
    wavelengths.append(wavelength)
    kappas.append(pars[0])

output_wb = Workbook()
output_ws = output_wb.active

output_ws.title = "Milk Kappa"

for i in range(len(wavelengths)):
    output_ws.append((wavelengths[i], kappas[i]))

output_wb.save('Output.xlsx')