import openpyxl as xl
from game_object_to_score_converter import (coral_point_converter, algae_point_converter,
                                            barge_point_converter, total_point_converter,
                                            coral_point_converter_auto)
from helpers import (team_assign, is_file_open)
from bar_graphs import (bar_graph)
from reefscapeScoutingAutomationFRC2025.score_average import barge_avg
from score_average import (game_piece_avg, total_points_avg, leave_avg)
from time import sleep
from sorting_commands import (teleop_data_sorter, auto_data_sorter)


filepath = '../data/Reefscape Scouting Sheet.xlsx'

while True:
    if is_file_open(filepath):
        print("File is open, waiting...")
        sleep(10)
    else:
        # Loads the workbook
        scouting_wb = xl.load_workbook(filepath)

        # Sorts unsorted raw data to a sorted sheet
        teleop_data_sorter(scouting_wb,'../data/Reefscape Scouting Sheet.xlsx', 'TeleOp - Unsorted Data', 'TeleOp - Raw Data')
        auto_data_sorter(scouting_wb,'../data/Reefscape Scouting Sheet.xlsx', 'Auto - Unsorted Data', 'Auto - Raw Data')

        # Assigns teams to proper cell
        team_assign(scouting_wb, 'TeleOp - Raw Data', 'TeleOp - Total Points')
        team_assign(scouting_wb, 'TeleOp - Raw Data', 'TeleOp - Average')
        team_assign(scouting_wb, 'Auto - Raw Data', 'Auto - Total Points')
        team_assign(scouting_wb, 'Auto - Raw Data', 'Auto - Average')

        # Point converter for teleop
        coral_point_converter(scouting_wb, 'TeleOp - Raw Data', 'TeleOp - Total Points')
        algae_point_converter(scouting_wb, 'TeleOp - Raw Data', 'TeleOp - Total Points')
        barge_point_converter(scouting_wb, 'TeleOp - Raw Data', 'TeleOp - Total Points')
        total_point_converter(scouting_wb, 'TeleOp - Total Points')

        # Point converter for auto
        coral_point_converter_auto(scouting_wb, 'Auto - Raw Data', 'Auto - Total Points')
        algae_point_converter(scouting_wb, 'Auto - Raw Data', 'Auto - Total Points')
        total_point_converter(scouting_wb, 'Auto - Total Points')

        # Calculates the average of teams for teleop
        game_piece_avg(scouting_wb, 'TeleOp - Raw Data', 'TeleOp - Average')
        barge_avg(scouting_wb, 'TeleOp - Raw Data', 'TeleOp - Average')
        total_points_avg(scouting_wb, 'TeleOp - Total Points', 'TeleOp - Average')

        # Calculates the average of teams for auto
        leave_avg(scouting_wb, 'Auto - Raw Data', 'Auto - Average')
        game_piece_avg(scouting_wb, 'Auto - Raw Data', 'Auto - Average')
        total_points_avg(scouting_wb, 'Auto - Total Points', 'Auto - Average')

        # Saves the changes in a workbook
        scouting_wb.save(filepath)
        print('Sheet Successfully Updated')
        sleep(10)