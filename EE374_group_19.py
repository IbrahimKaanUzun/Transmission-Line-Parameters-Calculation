import sys
import os
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QHBoxLayout, QComboBox, QLabel, QLineEdit, QFormLayout, QMessageBox
import math

# Conductor and tower type details with additional constraints
conductors = {
    'Hawk': {'Diameter': 21.793, 'GMR': 8.809, 'AC_resistance': 0.132, 'Current Capacity': 659},
    'Drake': {'Diameter': 28.143, 'GMR': 11.369, 'AC_resistance': 0.080, 'Current Capacity': 907},
    'Cardinal': {'Diameter': 30.378, 'GMR': 12.253, 'AC_resistance': 0.067, 'Current Capacity': 996},
    'Rail': {'Diameter': 29.591, 'GMR': 11.765, 'AC_resistance': 0.068, 'Current Capacity': 993},
    'Pheasant': {'Diameter': 35.103, 'GMR': 14.204, 'AC_resistance': 0.051, 'Current Capacity': 1187}
}

tower_types = {
    'Type-1: Narrow Base Tower': {'max_height': 39, 'min_height': 23, 'max_horizontal': 4, 'min_horizontal': 2.2, 'voltage': 66, 'max_conductors': 3},
    'Type-2: Single Circuit Delta Tower': {'max_height': 43, 'min_height': 38.25, 'max_horizontal': 11.5, 'min_horizontal': 9.4, 'max_horizontal_center': 8.9, 'voltage': 400, 'max_conductors': 4},
    'Type-3: Double Circuit Vertical Tower': {'max_height': 48.8, 'min_height': 36, 'max_horizontal': 5.35, 'min_horizontal': 1.8, 'voltage': 154, 'max_conductors': 3}
}

def check_constraints(tower_type, num_conductors, coordinates, num_circuits):
    if num_conductors > tower_types[tower_type]['max_conductors']:
        return "Number of conductors exceeds the maximum allowed for this tower type."
    
    if num_conductors <= 0:
        return "Number of conductors cannot be less than 1."
    
    if (tower_type == 'Type-3: Double Circuit Vertical Tower' and num_circuits not in [1, 2]) or (tower_type in ['Type-1: Narrow Base Tower', 'Type-2: Single Circuit Delta Tower'] and num_circuits != 1):
        return "Invalid number of circuits for the selected tower type."
    
    tower_specs = tower_types[tower_type]
    for circuit in coordinates:
        for phase, (x, y) in coordinates[circuit].items():
            if tower_type == 'Type-1: Narrow Base Tower':
                if not ((tower_specs['min_horizontal'] <= abs(x) <= tower_specs['max_horizontal']) and (x >= tower_specs['min_horizontal'] or x <= -tower_specs['min_horizontal']) and tower_specs['min_height'] <= y <= tower_specs['max_height']):
                    return f"Coordinates for Circuit {circuit} Phase {phase} are out of the allowed range for this tower type."
            elif tower_type == 'Type-2: Single Circuit Delta Tower':
                if phase == 'B':
                    if not (-8.9 <= x <= 8.9 and tower_specs['min_height'] <= y <= tower_specs['max_height']):
                        return f"Coordinates for Circuit {circuit} Phase {phase} are out of the allowed range for this tower type."
                else:
                    if not ((9.4 <= abs(x) <= 11.5) and tower_specs['min_height'] <= y <= tower_specs['max_height']):
                        return f"Coordinates for Circuit {circuit} Phase {phase} are out of the allowed range for this tower type."
            elif tower_type == 'Type-3: Double Circuit Vertical Tower':
                if not ((tower_specs['min_horizontal'] <= abs(x) <= tower_specs['max_horizontal']) and (x >= tower_specs['min_horizontal'] or x <= -tower_specs['min_horizontal']) and tower_specs['min_height'] <= y <= tower_specs['max_height']):
                    return f"Coordinates for Circuit {circuit} Phase {phase} are out of the allowed range for this tower type."
    return None  # No errors

class TransmissionLineCalc(QWidget):
    def __init__(self):
        super().__init__()
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle('Transmission Line Parameter Calculation Tool')
        
        main_layout = QHBoxLayout()  # Main layout to hold form and image
        form_layout = QFormLayout()  # Form layout for inputs and buttons

        self.tower_type_combo = QComboBox()
        self.tower_type_combo.addItems(list(tower_types.keys()))
        self.tower_type_combo.currentIndexChanged.connect(self.update_phase_inputs)
        self.tower_type_combo.currentIndexChanged.connect(self.update_image)
        form_layout.addRow('Tower Type:', self.tower_type_combo)

        self.num_circuits_input = QLineEdit()
        self.num_circuits_input.textChanged.connect(self.update_phase_inputs)
        form_layout.addRow('Number of Circuits:', self.num_circuits_input)

        self.phase_inputs = {}
        for circuit in range(1, 3):  # Support up to 2 circuits
            for phase in ['A', 'B', 'C']:
                coord_x = QLineEdit()
                coord_x.setPlaceholderText('x coordinate:')
                coord_y = QLineEdit()
                coord_y.setPlaceholderText('y coordinate:')
                self.phase_inputs[(circuit, phase)] = (coord_x, coord_y)
                phase_layout = QVBoxLayout()
                phase_layout.addWidget(QLabel(f'Circuit {circuit} Phase {phase} Coordinates (x,y):'))
                phase_layout.addWidget(coord_x)
                phase_layout.addWidget(coord_y)
                form_layout.addRow(phase_layout)

        self.num_conductors_input = QLineEdit()
        form_layout.addRow('Number of Conductors in Bundle:', self.num_conductors_input)

        self.distance_between_conductors_input = QLineEdit()
        form_layout.addRow('Distance Between Conductors in Bundle (cm):', self.distance_between_conductors_input)

        self.conductor_type_combo = QComboBox()
        self.conductor_type_combo.addItems(list(conductors.keys()))
        form_layout.addRow('Conductor Type:', self.conductor_type_combo)

        self.line_length_input = QLineEdit()
        form_layout.addRow('Transmission Line Length (km):', self.line_length_input)

        self.calculate_button = QPushButton('Calculate')
        self.calculate_button.setStyleSheet("background-color: blue; color: white;")
        self.calculate_button.clicked.connect(self.perform_calculation)
        form_layout.addRow(self.calculate_button)

        self.clear_button = QPushButton('Clear')
        self.clear_button.setStyleSheet("background-color: red; color: white;")
        self.clear_button.clicked.connect(self.clear_inputs)
        form_layout.addRow(self.clear_button)

        self.results_label = QLabel('')
        form_layout.addRow('Results:', self.results_label)

        main_layout.addLayout(form_layout)  # Add form layout to main layout

        # Add image to the right side
        self.image_label = QLabel()
        main_layout.addWidget(self.image_label)  # Add image label to main layout

        self.setLayout(main_layout)  # Set the main layout

        # Initial image
        self.update_image()

    def update_image(self):
        tower_type = self.tower_type_combo.currentText()
        current_directory = os.path.dirname(os.path.abspath(__file__))
        
        if tower_type == 'Type-1: Narrow Base Tower':
            image_path = os.path.join(current_directory, 'type1_image.jpg')
        elif tower_type == 'Type-2: Single Circuit Delta Tower':
            image_path = os.path.join(current_directory, 'type2_image.jpg')
        elif tower_type == 'Type-3: Double Circuit Vertical Tower':
            image_path = os.path.join(current_directory, 'type3_image.jpg')
        else:
            image_path = ''

        if image_path and os.path.exists(image_path):
            pixmap = QPixmap(image_path)
            self.image_label.setPixmap(pixmap)
            self.image_label.setScaledContents(True)
            self.image_label.setFixedSize(400, 400)  # Adjust size as needed
        else:
            self.image_label.setText('Image not found.')

    def update_phase_inputs(self):
        tower_type = self.tower_type_combo.currentText()
        num_circuits = int(self.num_circuits_input.text()) if self.num_circuits_input.text().isdigit() else 1
        
        for circuit in range(1, 3):
            for phase in ['A', 'B', 'C']:
                visible = (tower_type == 'Type-3: Double Circuit Vertical Tower' and num_circuits == 2) or circuit == 1
                self.phase_inputs[(circuit, phase)][0].setVisible(visible)
                self.phase_inputs[(circuit, phase)][1].setVisible(visible)

    def distance(self, x1, x2, y1, y2):
        return math.sqrt((x1 - x2) ** 2 + (y1 - y2) ** 2)
    
    def GMD_calculator(self, coordinates):
        ax, ay = coordinates[1]['A']
        bx, by = coordinates[1]['B']
        cx, cy = coordinates[1]['C']
        a2b = self.distance(ax, bx, ay, by)
        a2c = self.distance(ax, cx, ay, cy)
        b2c = self.distance(bx, cx, by, cy)
        return (a2b * a2c * b2c) ** (1 / 3)

    def GMR_calculator(self, GMR_conductor, num_conductors, distance_between_conductors):
        if num_conductors == 1:
            return GMR_conductor
        elif num_conductors == 2 :
            return ((GMR_conductor * distance_between_conductors) ** (1 / 2)) 
        elif num_conductors == 3 :
            return (((GMR_conductor * distance_between_conductors ** 2)) ** (1 / 3)) 
        elif num_conductors == 4:
            return ((GMR_conductor * (distance_between_conductors ** 2) * (distance_between_conductors * math.sqrt(2))) ** (1 / 4)) 
        
    def Req_calculator(self, r_eq, num_conductors, distance_between_conductors):
        if num_conductors == 1:
            return r_eq
        elif num_conductors == 2 :
            return ((r_eq * distance_between_conductors) ** (1 / 2)) 
        elif num_conductors == 3 :
            return ((r_eq * (distance_between_conductors ** 2)) ** (1 / 3)) 
        elif num_conductors == 4:
            return ((r_eq * (distance_between_conductors ** 2) * (distance_between_conductors * math.sqrt(2))) ** (1 / 4)) 

    def two_GMR(self, GMR, a1x, b1x, c1x, a1y, b1y, c1y, a2x, b2x, c2x, a2y, b2y, c2y):
        GMR1 = (self.distance(a1x, a2x, a1y, a2y) * GMR) ** (1/2)
        GMR2 = (self.distance(b1x, b2x, b1y, b2y) * GMR) ** (1/2)
        GMR3 = (self.distance(c1x, c2x, c1y, c2y) * GMR) ** (1/2)
        return (GMR1 * GMR2 * GMR3) ** (1/3)
    
    def two_Req(self, Req, a1x, b1x, c1x, a1y, b1y, c1y, a2x, b2x, c2x, a2y, b2y, c2y):
        Req1 = (self.distance(a1x, a2x, a1y, a2y) * Req) ** (1/2)
        Req2 = (self.distance(b1x, b2x, b1y, b2y) * Req) ** (1/2)
        Req3 = (self.distance(c1x, c2x, c1y, c2y) * Req) ** (1/2)
        return (Req1 * Req2 * Req3) ** (1/3)
        
    def two_GMD(self, a1x, b1x, c1x, a1y, b1y, c1y, a2x, b2x, c2x, a2y, b2y, c2y):
        a1b1 = self.distance(a1x, b1x, a1y, b1y)
        a1b2 = self.distance(a1x, b2x, a1y, b2y)
        a2b1 = self.distance(a2x, b1x, a2y, b1y)
        a2b2 = self.distance(a2x, b2x, a2y, b2y)
        a_b = (a1b1 * a1b2 * a2b1 * a2b2) ** (1/4)

        a1c1 = self.distance(a1x, c1x, a1y, c1y)
        a1c2 = self.distance(a1x, c2x, a1y, c2y)
        a2c1 = self.distance(a2x, c1x, a2y, c1y)
        a2c2 = self.distance(a2x, c2x, a2y, c2y)
        a_c = (a1c1 * a1c2 * a2c1 * a2c2) ** (1/4)

        b1c1 = self.distance(b1x, c1x, b1y, c1y)
        b1c2 = self.distance(b1x, c2x, b1y, c2y)
        b2c1 = self.distance(b2x, c1x, b2y, c1y)
        b2c2 = self.distance(b2x, c2x, b2y, c2y)
        b_c = (b1c1 * b1c2 * b2c1 * b2c2) ** (1/4)
        return (a_b * a_c * b_c) ** (1/3)

    def perform_calculation(self):
        # Check if all necessary inputs are provided
        if not self.num_circuits_input.text() or not self.num_conductors_input.text() or not self.distance_between_conductors_input.text() or not self.line_length_input.text():
            QMessageBox.critical(self, "Error", "All input fields must be filled.")
            return

        tower_type = self.tower_type_combo.currentText()
        num_conductors = int(self.num_conductors_input.text())
        num_circuits = int(self.num_circuits_input.text())

        coordinates = {}
        for circuit in range(1, num_circuits + 1):
            coordinates[circuit] = {}
            for phase in ['A', 'B', 'C']:
                x = float(self.phase_inputs[(circuit, phase)][0].text())
                y = float(self.phase_inputs[(circuit, phase)][1].text())
                coordinates[circuit][phase] = (x, y)

        all_coords = [coordinates[circuit][phase] for circuit in coordinates for phase in coordinates[circuit]]
        if len(all_coords) != len(set(all_coords)):
            QMessageBox.critical(self, "Error", "Two or more phases have the same coordinates.")
            return

        error_message = check_constraints(tower_type, num_conductors, coordinates, num_circuits)
        if error_message:
            QMessageBox.critical(self, "Error", error_message)
            return

        distance_between_conductors = float(self.distance_between_conductors_input.text()) / 100 #cm conversion
        conductor_type = self.conductor_type_combo.currentText()
        line_length = float(self.line_length_input.text())

        conductor_info = conductors[conductor_type]
        R = conductor_info['AC_resistance'] * line_length

        GMR_conductor = conductor_info['GMR'] / 1000  # Convert GMR to meters

        if tower_type == 'Type-3: Double Circuit Vertical Tower' and num_circuits == 2:
            a1x, a1y = coordinates[1]['A']
            b1x, b1y = coordinates[1]['B']
            c1x, c1y = coordinates[1]['C']
            a2x, a2y = coordinates[2]['A']
            b2x, b2y = coordinates[2]['B']
            c2x, c2y = coordinates[2]['C']

            GMR = self.two_GMR(GMR_conductor, a1x, b1x, c1x, a1y, b1y, c1y, a2x, b2x, c2x, a2y, b2y, c2y)
            Req = self.two_Req(conductor_info['Diameter'] / 2000, a1x, b1x, c1x, a1y, b1y, c1y, a2x, b2x, c2x, a2y, b2y, c2y)
            GMD = self.two_GMD(a1x, b1x, c1x, a1y, b1y, c1y, a2x, b2x, c2x, a2y, b2y, c2y)

            L = (2 * 10 ** -7 * math.log(GMD / GMR)) * line_length * 10**6  
            epsilon_0 = 8.854 * 10 ** -12  
            epsilon_r = 1  
            C = (2 * math.pi * epsilon_0 * epsilon_r / math.log(GMD / Req)) * line_length * 10**9  
            voltage = tower_types[tower_type]['voltage']
            current_capacity = conductors[conductor_type]['Current Capacity']
            MVA = 2 * voltage * current_capacity * math.sqrt(3) * 10 ** -3

            results = [
                f'Line Resistance: {R:.5f} Ω',
                f'Line Inductance: {L:.5f} mH',
                f'Line Capacitance: {C:.5f} µF',
                f'Line Capacity: {MVA:.5f} MVA',
              #  f'GMD: {GMD:.5f}',
              #  f'GMR: {GMR:.5f}',
              #  f'Req: {Req:.5f}'
            ]

        else:
            GMR = self.GMR_calculator(GMR_conductor, num_conductors, distance_between_conductors)  
            Req = self.Req_calculator(conductor_info['Diameter'] / 2000, num_conductors, distance_between_conductors)
            GMD = self.GMD_calculator(coordinates)  

            L = (2 * 10 ** -7 * math.log(GMD / GMR)) * line_length * 10**6  
            epsilon_0 = 8.854 * 10 ** -12  
            epsilon_r = 1  
            C = (2 * math.pi * epsilon_0 * epsilon_r / math.log(GMD / Req)) * line_length * 10**9  
            voltage = tower_types[tower_type]['voltage']
            current_capacity = conductors[conductor_type]['Current Capacity']
            MVA = voltage * current_capacity * math.sqrt(3) * 10 ** -3

            results = [
                f'Line Resistance: {R:.5f} Ω',
                f'Line Inductance: {L:.5f} mH',
                f'Line Capacitance: {C:.5f} µF',
                f'Line Capacity: {MVA:.5f} MVA',
              #  f'GMD: {GMD:.5f}',
              #  f'GMR: {GMR:.5f}',
              #  f'Req: {Req:.5f}'
            ]
        self.results_label.setText("\n".join(results))

    def clear_inputs(self):
        self.num_circuits_input.clear()
        for circuit in range(1, 3):
            for phase in ['A', 'B', 'C']:
                self.phase_inputs[(circuit, phase)][0].clear()
                self.phase_inputs[(circuit, phase)][1].clear()
        self.num_conductors_input.clear()
        self.distance_between_conductors_input.clear()
        self.line_length_input.clear()
        self.results_label.clear()

app = QApplication(sys.argv)
window = TransmissionLineCalc()
window.show()
sys.exit(app.exec())
