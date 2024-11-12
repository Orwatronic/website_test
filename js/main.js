// Donation Button Functionality
document.addEventListener('DOMContentLoaded', function() {
    const donationButtons = document.querySelectorAll('.amount-btn');
    const donateNowButton = document.querySelector('.donation-options + .button.primary');
    let selectedAmount = '';

    // Handle donation amount selection
    donationButtons.forEach(button => {
        button.addEventListener('click', function() {
            // Remove active class from all buttons
            donationButtons.forEach(btn => btn.classList.remove('active'));
            
            // Add active class to clicked button
            this.classList.add('active');
            
            // Store selected amount
            selectedAmount = this.textContent === 'Other' ? 'custom' : this.textContent;
            
            // Update donate button text
            donateNowButton.textContent = selectedAmount === 'custom' 
                ? 'Enter Amount' 
                : `Donate ${selectedAmount}`;
        });
    });

    // Handle "Other" amount
    const otherButton = document.querySelector('.amount-btn:last-child');
    otherButton.addEventListener('click', function() {
        const amount = prompt('Enter custom amount in USD:');
        if (amount && !isNaN(amount)) {
            selectedAmount = `$${parseFloat(amount).toFixed(2)}`;
            donateNowButton.textContent = `Donate ${selectedAmount}`;
        }
    });

    // Handle donation button click
    donateNowButton.addEventListener('click', function(e) {
        e.preventDefault();
        if (!selectedAmount) {
            alert('Please select an amount first');
            return;
        }
        // Here you would typically redirect to a payment processor
        alert(`Processing donation of ${selectedAmount}. In production, this would connect to a payment processor.`);
    });
});

// Contact Form Handling
document.querySelector('.contact-form').addEventListener('submit', async function(e) {
    e.preventDefault();

    // Get form data
    const formData = {
        name: document.getElementById('name').value,
        email: document.getElementById('email').value,
        subject: document.getElementById('subject').value,
        message: document.getElementById('message').value
    };

    // Validate form data
    if (!validateForm(formData)) {
        return;
    }

    // Show loading state
    const submitButton = this.querySelector('button[type="submit"]');
    const originalButtonText = submitButton.textContent;
    submitButton.textContent = 'Sending...';
    submitButton.disabled = true;

    try {
        // In production, replace this with your actual API endpoint
        // await fetch('/api/contact', {
        //     method: 'POST',
        //     headers: {
        //         'Content-Type': 'application/json',
        //     },
        //     body: JSON.stringify(formData)
        // });

        // Simulate API call
        await new Promise(resolve => setTimeout(resolve, 1000));

        // Show success message
        showMessage('Message sent successfully!', 'success');
        this.reset();

    } catch (error) {
        showMessage('Failed to send message. Please try again.', 'error');
    } finally {
        // Reset button state
        submitButton.textContent = originalButtonText;
        submitButton.disabled = false;
    }
});

// Form validation
function validateForm(data) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    
    if (data.name.length < 2) {
        showMessage('Please enter a valid name', 'error');
        return false;
    }
    
    if (!emailRegex.test(data.email)) {
        showMessage('Please enter a valid email address', 'error');
        return false;
    }
    
    if (data.message.length < 10) {
        showMessage('Message must be at least 10 characters long', 'error');
        return false;
    }
    
    return true;
}

// Message display functionality
function showMessage(message, type) {
    // Create message element if it doesn't exist
    let messageDiv = document.querySelector('.message-div');
    if (!messageDiv) {
        messageDiv = document.createElement('div');
        messageDiv.className = 'message-div';
        document.querySelector('.contact-form').insertBefore(
            messageDiv,
            document.querySelector('.contact-form button')
        );
    }

    // Set message content and style
    messageDiv.textContent = message;
    messageDiv.className = `message-div ${type}`;

    // Remove message after 5 seconds
    setTimeout(() => {
        messageDiv.remove();
    }, 5000);
} 