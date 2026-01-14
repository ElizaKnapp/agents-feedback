// Script

// Initialize projects array from localStorage or use empty array
let projects = JSON.parse(localStorage.getItem('projects')) || [];

// DOM elements
const projectNameSelect = document.getElementById('projectName');
const addProjectForm = document.getElementById('addProjectForm');
const newProjectNameInput = document.getElementById('newProjectName');
const projectDescriptionInput = document.getElementById('projectDescription');

// Function to populate the dropdown with projects
function populateDropdown() {
    // Clear existing options except the first one
    projectNameSelect.innerHTML = '<option value="">-- Select a project --</option>';
    
    // Add each project as an option
    projects.forEach((project, index) => {
        const option = document.createElement('option');
        option.value = project.name;
        option.textContent = project.name;
        option.setAttribute('data-description', project.description);
        projectNameSelect.appendChild(option);
    });
}

// Function to add a new project
function addProject(name, description) {
    const newProject = {
        name: name.trim(),
        description: description.trim()
    };
    
    // Check if project name already exists
    if (projects.some(p => p.name.toLowerCase() === newProject.name.toLowerCase())) {
        alert('A project with this name already exists!');
        return false;
    }
    
    projects.push(newProject);
    localStorage.setItem('projects', JSON.stringify(projects));
    populateDropdown();
    return true;
}

// Handle form submission for adding new projects
addProjectForm.addEventListener('submit', (e) => {
    e.preventDefault();
    
    const name = newProjectNameInput.value;
    const description = projectDescriptionInput.value;
    
    if (addProject(name, description)) {
        // Clear the form
        newProjectNameInput.value = '';
        projectDescriptionInput.value = '';
    }
});

// Populate dropdown on page load
populateDropdown();
