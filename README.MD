# Speech-to-Email Generator

This project is a Python script that converts spoken input into an email response using Google Cloud Speech-to-Text for speech recognition and OpenAI's GPT-3.5 for generating the email content. It integrates with Microsoft Outlook for email handling.

## Setup

### Prerequisites

- Python 3.x installed on your system
- Google Cloud account with Speech-to-Text API enabled
- OpenAI API key
- Microsoft Outlook installed on your system (for email handling)

### Installation

1. Clone this repository to your local machine.

2. Install the required Python packages using pip:

   ```bash
   pip install -r requirements.txt
   ```

3. Create an `.env` file in the root directory of the project and add your OpenAI API key:

   ```plaintext
   OPENAI_API_KEY=your_openai_api_key_here
   ```

4. Place your Google Cloud service account credentials in a file named `credentials.json` in the root directory of the project.

5. Ensure that Microsoft Outlook is configured and running on your system.

## Usage

1. Run the Python script `app.py`.

2. Select an email in Microsoft Outlook to reply to, or start a new email.

3. Speak your email response when prompted. Press Enter when finished speaking.

4. The script will generate an email response based on your speech input and open a draft in Microsoft Outlook for you to review and send.

## License

All rights reserved. No part of this project may be reproduced, distributed, or transmitted in any form or by any means, including photocopying, recording, or other electronic or mechanical methods, without the prior written permission of the author, except in the case of brief quotations embodied in critical reviews and certain other noncommercial uses permitted by copyright law.

## Screenshots
