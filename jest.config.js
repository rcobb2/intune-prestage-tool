export default {
  // Enable experimental VM modules for dynamic import of ES modules
  nodeOptions: '--experimental-vm-modules',
  preset: 'ts-jest',
  testEnvironment: 'node',
  transform: {
    '^.+\\.tsx?$': 'ts-jest',
  },
  moduleFileExtensions: ['ts', 'js', 'json', 'node'],
  // Adjust the root directory if needed

};
