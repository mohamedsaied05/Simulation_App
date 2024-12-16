import 'package:flutter/material.dart';
import 'package:simulation_app/screens/parallel_server_simulation_screen.dart';
import 'package:simulation_app/screens/probability_simulation_screen.dart';
import 'package:simulation_app/screens/static_simulation_screen.dart';

void main() {
  runApp(const MyApp());
}

class MyApp extends StatelessWidget {
  const MyApp({super.key});

  @override
  Widget build(BuildContext context) {
    return const MaterialApp(
      debugShowCheckedModeBanner: false,
      home: WelcomeScreen(),
    );
  }
}

class WelcomeScreen extends StatelessWidget {
  const WelcomeScreen({super.key});

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      body: Container(
        decoration: const BoxDecoration(
          gradient: LinearGradient(
            colors: [Colors.tealAccent, Colors.blueGrey],
            begin: Alignment.topCenter,
            end: Alignment.bottomCenter,
          ),
        ),
        child: Center(
          child: Padding(
            padding: const EdgeInsets.all(20.0),
            child: Container(
              margin: const EdgeInsets.only(top: 100),
              child: Column(
                mainAxisAlignment: MainAxisAlignment.start,
                children: <Widget>[
                  const Text(
                    'Welcome to our Simulation App!',
                    style: TextStyle(
                      fontSize: 24,
                      fontWeight: FontWeight.bold,
                      color: Colors.white,
                      shadows: [
                        Shadow(
                          blurRadius: 5.0,
                          color: Colors.black54,
                          offset: Offset(2.0, 2.0),
                        ),
                      ],
                    ),
                    textAlign: TextAlign.center,
                  ),
                  const SizedBox(height: 30),
                  const Text(
                    'Choose Simulation Type:',
                    style: TextStyle(
                      fontSize: 20,
                      fontWeight: FontWeight.w600,
                      color: Colors.white,
                      shadows: [
                        Shadow(
                          blurRadius: 5.0,
                          color: Colors.black54,
                          offset: Offset(2.0, 2.0),
                        ),
                      ],
                    ),
                    textAlign: TextAlign.center,
                  ),
                  const SizedBox(height: 30),
                  // First Button with Icon and Animation
                  AnimatedButton(
                    label: 'Static',
                    onPressed: () {
                      Navigator.push(
                        context,
                        MaterialPageRoute(
                          builder: (context) => const ExcelSimulationScreen(),
                        ),
                      );
                    },
                    color: Colors.teal,
                    icon: Icons.scatter_plot,
                  ),
                  const SizedBox(height: 20),
                  // Second Button with Icon and Animation
                  AnimatedButton(
                    label: 'Probability',
                    onPressed: () {
                      Navigator.push(
                        context,
                        MaterialPageRoute(
                          builder: (context) =>
                              const ProbabilitySimulationScreen(),
                        ),
                      );
                    },
                    color: Colors.teal[700]!,
                    icon: Icons.calculate,
                  ),
                  const SizedBox(height: 20),
                  // Third Button with Icon and Animation
                  AnimatedButton(
                    label: 'Parallel Server',
                    onPressed: () {
                      Navigator.push(
                        context,
                        MaterialPageRoute(
                          builder: (context) =>
                              const ParallelServerSimulationScreen(),
                        ),
                      );
                    },
                    color: Colors.teal[900]!,
                    icon: Icons.align_horizontal_left_outlined,
                  ),
                ],
              ),
            ),
          ),
        ),
      ),
    );
  }
}

// Custom Animated Button with Icon
class AnimatedButton extends StatelessWidget {
  final String label;
  final VoidCallback onPressed;
  final Color color;
  final IconData icon;

  const AnimatedButton({
    super.key,
    required this.label,
    required this.onPressed,
    required this.color,
    required this.icon,
  });

  @override
  Widget build(BuildContext context) {
    return InkWell(
      onTap: onPressed,
      splashColor: color.withOpacity(0.5),
      borderRadius: BorderRadius.circular(30),
      child: Container(
        width: 300,
        padding: const EdgeInsets.symmetric(vertical: 15, horizontal: 30),
        decoration: BoxDecoration(
          color: color,
          borderRadius: BorderRadius.circular(30),
          boxShadow: [
            BoxShadow(
              color: color.withOpacity(0.6),
              blurRadius: 5,
              offset: const Offset(0, 4),
            ),
          ],
        ),
        child: Row(
          mainAxisAlignment: MainAxisAlignment.center,
          children: [
            Icon(icon, color: Colors.white, size: 24),
            const SizedBox(width: 10),
            Text(
              label,
              style: const TextStyle(fontSize: 18, color: Colors.white),
            ),
          ],
        ),
      ),
    );
  }
}
