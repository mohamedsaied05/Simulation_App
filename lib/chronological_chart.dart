import 'package:flutter/material.dart';
import 'package:fl_chart/fl_chart.dart';

class EventsChart extends StatelessWidget {
  final List<Map<String, dynamic>> events;

  const EventsChart({super.key, required this.events});

  @override
  Widget build(BuildContext context) {
    // Separate arrival and departure events
    final arrivalEvents = events
        .where((event) => event['Event_Type'] == 'Arrival')
        .map((event) => event['Clock_Time'] as int)
        .toList();

    final departureEvents = events
        .where((event) => event['Event_Type'] == 'Departure')
        .map((event) => event['Clock_Time'] as int)
        .toList();

    return Scaffold(
      appBar: AppBar(title: const Text('Events Chart')),
      body: Padding(
        padding: const EdgeInsets.all(16.0),
        child: BarChart(
          BarChartData(
            alignment: BarChartAlignment.spaceAround,
            maxY: (events
                        .map((e) => e['Clock_Time'] as int)
                        .reduce((a, b) => a > b ? a : b) *
                    1.2)
                .toDouble(),
            barGroups: _createBarGroups(arrivalEvents, departureEvents),
            barTouchData: BarTouchData(enabled: true),
            titlesData: FlTitlesData(
              leftTitles: AxisTitles(
                sideTitles: SideTitles(
                  showTitles: true,
                  getTitlesWidget: (value, _) => Text(value.toInt().toString()),
                ),
              ),
              bottomTitles: AxisTitles(
                sideTitles: SideTitles(
                  showTitles: true,
                  getTitlesWidget: (value, _) => Text('T${value.toInt()}'),
                ),
              ),
            ),
          ),
        ),
      ),
    );
  }

  List<BarChartGroupData> _createBarGroups(
      List<int> arrivalEvents, List<int> departureEvents) {
    // Define bar groups for arrival and departure events
    List<BarChartGroupData> barGroups = [];

    for (int i = 0; i < arrivalEvents.length; i++) {
      barGroups.add(
        BarChartGroupData(
          x: i,
          barRods: [
            BarChartRodData(
              toY: arrivalEvents[i].toDouble(),
              color: Colors.blue, // Arrival bars in blue
              width: 8,
            ),
            BarChartRodData(
              toY: departureEvents[i].toDouble(),
              color: Colors.red, // Departure bars in red
              width: 8,
            ),
          ],
          showingTooltipIndicators: [0, 1],
        ),
      );
    }

    return barGroups;
  }
}
