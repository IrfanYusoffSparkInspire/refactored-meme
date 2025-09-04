const express = require('express');
const multer = require('multer');
const path = require('path');
const fs = require('fs').promises;
const { spawn } = require('child_process');
const cors = require('cors');

const app = express();
const PORT = process.env.PORT || 3000;

// Middleware
app.use(cors());
app.use(express.json({ limit: '50mb' }));
app.use(express.urlencoded({ limit: '50mb', extended: true }));
app.use(express.static('.'));

// Create necessary directories
const createDirectories = async () => {
    const dirs = ['temp_images', 'generated_proposals', 'uploads'];
    for (const dir of dirs) {
        try {
            await fs.mkdir(dir, { recursive: true });
        } catch (error) {
            console.error(`Error creating directory ${dir}:`, error);
        }
    }
};

// Configure multer for file uploads
const storage = multer.diskStorage({
    destination: async (req, file, cb) => {
        await createDirectories();
        cb(null, 'temp_images/');
    },
    filename: (req, file, cb) => {
        const uniqueSuffix = Date.now() + '-' + Math.round(Math.random() * 1E9);
        cb(null, file.fieldname + '-' + uniqueSuffix + path.extname(file.originalname));
    }
});

const upload = multer({ 
    storage: storage,
    limits: {
        fileSize: 10 * 1024 * 1024 // 10MB limit
    }
});

// Serve the main HTML file
app.get('/', (req, res) => {
    res.sendFile(path.join(__dirname, 'index.html'));
});

// Generate proposal endpoint
app.post('/api/generate-proposal', upload.single('image'), async (req, res) => {
    try {
        console.log('📝 Processing proposal generation request...');
        
        // Get form data from request body
        const {
            building_name = '',
            address = '',
            survey_date = '',
            prepared_by = '',
            prepared_date = '',
            type_building = '',
            building_manager_name = '',
            building_manager_email = '',
            building_manager_phone = '',
            building_manager_company = '',
            otic = '',
            tap_new_or_spare = '',
            tapping_location = '',
            tapping_location_level = '',
            site_assessment_mccb = '',
            tnb_meter = '',
            tnb_na = '',
            parking_location = '',
            ev_charger_model = '',
            network_strength = '',
            imageData = '', // Base64 image data from canvas (MSB)
            mccbImageData = '', // Base64 image data for MCCB
            tpsldImageData = '', // Base64 image data for TP_SLD
            tpmccbcompartmentImageData = '', // Base64 image data for TP_MCCB_COMPARTMENT
            tptappinglocImageData = '', // Base64 image data for TP_TAPPING_LOC
            tprouting1ImageData = '', // Base64 image data for TP_ROUTING_1
            tprouting2ImageData = '', // Base64 image data for TP_ROUTING_2
            tprouting3ImageData = '' // Base64 image data for TP_ROUTING_3
        } = req.body;

        console.log('📋 Form data received:', { building_name, address: address.substring(0, 50) });
        
        // Validate required template file
        const templatePath = path.join(__dirname, 'TP_Template.pptx');
        try {
            await fs.access(templatePath);
            console.log('✅ Template file found:', templatePath);
        } catch (error) {
            console.error('❌ Template file not found:', templatePath);
            return res.status(400).json({
                success: false,
                message: 'Template file TP_Template.pptx not found. Please ensure it exists in the project directory.'
            });
        }

        // Prepare data for Python script
        const proposalData = {
            building_name,
            address,
            survey_date,
            prepared_by,
            prepared_date,
            type_building,
            building_manager_name,
            building_manager_email,
            building_manager_phone,
            building_manager_company,
            otic,
            tap_new_or_spare,
            tapping_location,
            tapping_location_level,
            site_assessment_mccb,
            tnb_meter,
            tnb_na,
            parking_location,
            noofchargers: '2', // Fixed to 2 chargers
            ev_charger_model,
            network_strength
        };

        // Save image data if provided
        let msbImagePath = null;
        let mccbImagePath = null;
        let tpsldImagePath = null;
        let tpmccbcompartmentImagePath = null;
        let tptappinglocImagePath = null;
        let tprouting1ImagePath = null;
        let tprouting2ImagePath = null;
        let tprouting3ImagePath = null;
        
        // Process MSB image
        if (imageData && imageData.startsWith('data:image/')) {
            try {
                console.log('🖼️ Processing MSB image data...');
                console.log('📊 MSB Image data length:', imageData.length);
                console.log('📊 MSB Image data prefix:', imageData.substring(0, 50));
                
                // Extract base64 data
                const base64Data = imageData.split(',')[1];
                const imageBuffer = Buffer.from(base64Data, 'base64');
                console.log('📊 MSB Image buffer size:', imageBuffer.length, 'bytes');
                
                // Save to temp file
                const timestamp = Date.now();
                msbImagePath = path.join(__dirname, 'temp_images', `msb_image_${timestamp}.png`);
                await fs.writeFile(msbImagePath, imageBuffer);
                console.log('✅ MSB Image saved to:', msbImagePath);
                
                // Verify file was saved
                const stats = await fs.stat(msbImagePath);
                console.log('📁 Saved MSB image file size:', stats.size, 'bytes');
                
            } catch (error) {
                console.error('❌ Error processing MSB image:', error);
                return res.status(400).json({
                    success: false,
                    message: 'Error processing MSB image data'
                });
            }
        } else {
            console.log('⚠️ No valid MSB image data provided');
        }
        
        // Process MCCB image
        if (mccbImageData && mccbImageData.startsWith('data:image/')) {
            try {
                console.log('🖼️ Processing MCCB image data...');
                console.log('📊 MCCB Image data length:', mccbImageData.length);
                console.log('📊 MCCB Image data prefix:', mccbImageData.substring(0, 50));
                
                // Extract base64 data
                const base64Data = mccbImageData.split(',')[1];
                const imageBuffer = Buffer.from(base64Data, 'base64');
                console.log('📊 MCCB Image buffer size:', imageBuffer.length, 'bytes');
                
                // Save to temp file
                const timestamp = Date.now();
                mccbImagePath = path.join(__dirname, 'temp_images', `mccb_image_${timestamp}.png`);
                await fs.writeFile(mccbImagePath, imageBuffer);
                console.log('✅ MCCB Image saved to:', mccbImagePath);
                
                // Verify file was saved
                const stats = await fs.stat(mccbImagePath);
                console.log('📁 Saved MCCB image file size:', stats.size, 'bytes');
                
            } catch (error) {
                console.error('❌ Error processing MCCB image:', error);
                return res.status(400).json({
                    success: false,
                    message: 'Error processing MCCB image data'
                });
            }
        } else {
            console.log('⚠️ No valid MCCB image data provided');
        }
        
        // Process TP_SLD image
        if (tpsldImageData && tpsldImageData.startsWith('data:image/')) {
            try {
                console.log('🖼️ Processing TP_SLD image data...');
                console.log('📊 TP_SLD Image data length:', tpsldImageData.length);
                console.log('📊 TP_SLD Image data prefix:', tpsldImageData.substring(0, 50));
                
                // Extract base64 data
                const base64Data = tpsldImageData.split(',')[1];
                const imageBuffer = Buffer.from(base64Data, 'base64');
                console.log('📊 TP_SLD Image buffer size:', imageBuffer.length, 'bytes');
                
                // Save to temp file
                const timestamp = Date.now();
                tpsldImagePath = path.join(__dirname, 'temp_images', `tpsld_image_${timestamp}.png`);
                await fs.writeFile(tpsldImagePath, imageBuffer);
                console.log('✅ TP_SLD Image saved to:', tpsldImagePath);
                
                // Verify file was saved
                const stats = await fs.stat(tpsldImagePath);
                console.log('📁 Saved TP_SLD image file size:', stats.size, 'bytes');
                
            } catch (error) {
                console.error('❌ Error processing TP_SLD image:', error);
                return res.status(400).json({
                    success: false,
                    message: 'Error processing TP_SLD image data'
                });
            }
        } else {
            console.log('⚠️ No valid TP_SLD image data provided');
        }
        
        // Process TP_MCCB_COMPARTMENT image
        if (tpmccbcompartmentImageData && tpmccbcompartmentImageData.startsWith('data:image/')) {
            try {
                console.log('🖼️ Processing TP_MCCB_COMPARTMENT image data...');
                console.log('📊 TP_MCCB_COMPARTMENT Image data length:', tpmccbcompartmentImageData.length);
                console.log('📊 TP_MCCB_COMPARTMENT Image data prefix:', tpmccbcompartmentImageData.substring(0, 50));
                
                // Extract base64 data
                const base64Data = tpmccbcompartmentImageData.split(',')[1];
                const imageBuffer = Buffer.from(base64Data, 'base64');
                console.log('📊 TP_MCCB_COMPARTMENT Image buffer size:', imageBuffer.length, 'bytes');
                
                // Save to temp file
                const timestamp = Date.now();
                tpmccbcompartmentImagePath = path.join(__dirname, 'temp_images', `tpmccbcompartment_image_${timestamp}.png`);
                await fs.writeFile(tpmccbcompartmentImagePath, imageBuffer);
                console.log('✅ TP_MCCB_COMPARTMENT Image saved to:', tpmccbcompartmentImagePath);
                
                // Verify file was saved
                const stats = await fs.stat(tpmccbcompartmentImagePath);
                console.log('📁 Saved TP_MCCB_COMPARTMENT image file size:', stats.size, 'bytes');
                
            } catch (error) {
                console.error('❌ Error processing TP_MCCB_COMPARTMENT image:', error);
                return res.status(400).json({
                    success: false,
                    message: 'Error processing TP_MCCB_COMPARTMENT image data'
                });
            }
        } else {
            console.log('⚠️ No valid TP_MCCB_COMPARTMENT image data provided');
        }
        
        // Process TP_TAPPING_LOC image
        if (tptappinglocImageData && tptappinglocImageData.startsWith('data:image/')) {
            try {
                console.log('🖼️ Processing TP_TAPPING_LOC image data...');
                console.log('📊 TP_TAPPING_LOC Image data length:', tptappinglocImageData.length);
                console.log('📊 TP_TAPPING_LOC Image data prefix:', tptappinglocImageData.substring(0, 50));
                
                // Extract base64 data
                const base64Data = tptappinglocImageData.split(',')[1];
                const imageBuffer = Buffer.from(base64Data, 'base64');
                console.log('📊 TP_TAPPING_LOC Image buffer size:', imageBuffer.length, 'bytes');
                
                // Save to temp file
                const timestamp = Date.now();
                tptappinglocImagePath = path.join(__dirname, 'temp_images', `tptappingloc_image_${timestamp}.png`);
                await fs.writeFile(tptappinglocImagePath, imageBuffer);
                console.log('✅ TP_TAPPING_LOC Image saved to:', tptappinglocImagePath);
                
                // Verify file was saved
                const stats = await fs.stat(tptappinglocImagePath);
                console.log('📁 Saved TP_TAPPING_LOC image file size:', stats.size, 'bytes');
                
            } catch (error) {
                console.error('❌ Error processing TP_TAPPING_LOC image:', error);
                return res.status(400).json({
                    success: false,
                    message: 'Error processing TP_TAPPING_LOC image data'
                });
            }
        } else {
            console.log('⚠️ No valid TP_TAPPING_LOC image data provided');
        }

        // Process TP_ROUTING_1 image
        if (tprouting1ImageData && tprouting1ImageData.startsWith('data:image/')) {
            try {
                console.log('🖼️ Processing TP_ROUTING_1 image data...');
                console.log('📊 TP_ROUTING_1 Image data length:', tprouting1ImageData.length);
                console.log('📊 TP_ROUTING_1 Image data prefix:', tprouting1ImageData.substring(0, 50));
                
                // Extract base64 data
                const base64Data = tprouting1ImageData.split(',')[1];
                const imageBuffer = Buffer.from(base64Data, 'base64');
                console.log('📊 TP_ROUTING_1 Image buffer size:', imageBuffer.length, 'bytes');
                
                // Save to temp file
                const timestamp = Date.now();
                tprouting1ImagePath = path.join(__dirname, 'temp_images', `tprouting1_image_${timestamp}.png`);
                await fs.writeFile(tprouting1ImagePath, imageBuffer);
                console.log('✅ TP_ROUTING_1 Image saved to:', tprouting1ImagePath);
                
                // Verify file was saved
                const stats = await fs.stat(tprouting1ImagePath);
                console.log('📁 Saved TP_ROUTING_1 image file size:', stats.size, 'bytes');
                
            } catch (error) {
                console.error('❌ Error processing TP_ROUTING_1 image:', error);
                return res.status(400).json({
                    success: false,
                    message: 'Error processing TP_ROUTING_1 image data'
                });
            }
        } else {
            console.log('⚠️ No valid TP_ROUTING_1 image data provided');
        }

        // Process TP_ROUTING_2 image
        if (tprouting2ImageData && tprouting2ImageData.startsWith('data:image/')) {
            try {
                console.log('🖼️ Processing TP_ROUTING_2 image data...');
                console.log('📊 TP_ROUTING_2 Image data length:', tprouting2ImageData.length);
                console.log('📊 TP_ROUTING_2 Image data prefix:', tprouting2ImageData.substring(0, 50));
                
                // Extract base64 data
                const base64Data = tprouting2ImageData.split(',')[1];
                const imageBuffer = Buffer.from(base64Data, 'base64');
                console.log('📊 TP_ROUTING_2 Image buffer size:', imageBuffer.length, 'bytes');
                
                // Save to temp file
                const timestamp = Date.now();
                tprouting2ImagePath = path.join(__dirname, 'temp_images', `tprouting2_image_${timestamp}.png`);
                await fs.writeFile(tprouting2ImagePath, imageBuffer);
                console.log('✅ TP_ROUTING_2 Image saved to:', tprouting2ImagePath);
                
                // Verify file was saved
                const stats = await fs.stat(tprouting2ImagePath);
                console.log('📁 Saved TP_ROUTING_2 image file size:', stats.size, 'bytes');
                
            } catch (error) {
                console.error('❌ Error processing TP_ROUTING_2 image:', error);
                return res.status(400).json({
                    success: false,
                    message: 'Error processing TP_ROUTING_2 image data'
                });
            }
        } else {
            console.log('⚠️ No valid TP_ROUTING_2 image data provided');
        }

        // Process TP_ROUTING_3 image
        if (tprouting3ImageData && tprouting3ImageData.startsWith('data:image/')) {
            try {
                console.log('🖼️ Processing TP_ROUTING_3 image data...');
                console.log('📊 TP_ROUTING_3 Image data length:', tprouting3ImageData.length);
                console.log('📊 TP_ROUTING_3 Image data prefix:', tprouting3ImageData.substring(0, 50));
                
                // Extract base64 data
                const base64Data = tprouting3ImageData.split(',')[1];
                const imageBuffer = Buffer.from(base64Data, 'base64');
                console.log('📊 TP_ROUTING_3 Image buffer size:', imageBuffer.length, 'bytes');
                
                // Save to temp file
                const timestamp = Date.now();
                tprouting3ImagePath = path.join(__dirname, 'temp_images', `tprouting3_image_${timestamp}.png`);
                await fs.writeFile(tprouting3ImagePath, imageBuffer);
                console.log('✅ TP_ROUTING_3 Image saved to:', tprouting3ImagePath);
                
                // Verify file was saved
                const stats = await fs.stat(tprouting3ImagePath);
                console.log('📁 Saved TP_ROUTING_3 image file size:', stats.size, 'bytes');
                
            } catch (error) {
                console.error('❌ Error processing TP_ROUTING_3 image:', error);
                return res.status(400).json({
                    success: false,
                    message: 'Error processing TP_ROUTING_3 image data'
                });
            }
        } else {
            console.log('⚠️ No valid TP_ROUTING_3 image data provided');
        }

        // Generate unique output filename
        const timestamp = new Date().toISOString().replace(/[:.]/g, '-');
        const safeClientName = building_name.replace(/[^a-zA-Z0-9]/g, '_').substring(0, 20) || 'proposal';
        const outputFilename = `${safeClientName}_${timestamp}.pptx`;
        const outputPath = path.join(__dirname, 'generated_proposals', outputFilename);

        // Call Python script to generate PowerPoint
        console.log('🐍 Calling Python script to generate proposal...');
        
        const pythonArgs = [
            'proposal_processor.py',
            '--template', templatePath,
            '--output', outputPath,
            '--data', JSON.stringify(proposalData)
        ];

        if (msbImagePath) {
            pythonArgs.push('--msb-image', msbImagePath);
        }
        
        if (mccbImagePath) {
            pythonArgs.push('--mccb-image', mccbImagePath);
        }
        
        if (tpsldImagePath) {
            pythonArgs.push('--tpsld-image', tpsldImagePath);
        }
        
        if (tpmccbcompartmentImagePath) {
            pythonArgs.push('--tpmccbcompartment-image', tpmccbcompartmentImagePath);
        }
        
        if (tptappinglocImagePath) {
            pythonArgs.push('--tptappingloc-image', tptappinglocImagePath);
        }
        
        if (tprouting1ImagePath) {
            pythonArgs.push('--tprouting1-image', tprouting1ImagePath);
        }
        
        if (tprouting2ImagePath) {
            pythonArgs.push('--tprouting2-image', tprouting2ImagePath);
        }
        
        if (tprouting3ImagePath) {
            pythonArgs.push('--tprouting3-image', tprouting3ImagePath);
        }

        const pythonProcess = spawn('python3', pythonArgs, {
            cwd: __dirname,
            stdio: ['pipe', 'pipe', 'pipe']
        });

        let stdout = '';
        let stderr = '';

        pythonProcess.stdout.on('data', (data) => {
            stdout += data.toString();
        });

        pythonProcess.stderr.on('data', (data) => {
            stderr += data.toString();
        });

        pythonProcess.on('close', async (code) => {
            console.log(`🐍 Python process completed with code: ${code}`);
            console.log('📝 Python stdout output:');
            console.log(stdout);
            if (stderr) {
                console.log('⚠️ Python stderr output:');
                console.log(stderr);
            }
            
            try {
                // Clean up temporary image files
                if (msbImagePath) {
                    try {
                        await fs.unlink(msbImagePath);
                        console.log('🗑️ Cleaned up temporary MSB image file');
                    } catch (cleanupError) {
                        console.warn('⚠️ Could not clean up temporary MSB image:', cleanupError);
                    }
                }
                
                if (mccbImagePath) {
                    try {
                        await fs.unlink(mccbImagePath);
                        console.log('🗑️ Cleaned up temporary MCCB image file');
                    } catch (cleanupError) {
                        console.warn('⚠️ Could not clean up temporary MCCB image:', cleanupError);
                    }
                }
                
                if (tpsldImagePath) {
                    try {
                        await fs.unlink(tpsldImagePath);
                        console.log('🗑️ Cleaned up temporary TP_SLD image file');
                    } catch (cleanupError) {
                        console.warn('⚠️ Could not clean up temporary TP_SLD image:', cleanupError);
                    }
                }
                
                if (tpmccbcompartmentImagePath) {
                    try {
                        await fs.unlink(tpmccbcompartmentImagePath);
                        console.log('🗑️ Cleaned up temporary TP_MCCB_COMPARTMENT image file');
                    } catch (cleanupError) {
                        console.warn('⚠️ Could not clean up temporary TP_MCCB_COMPARTMENT image:', cleanupError);
                    }
                }
                
                if (tptappinglocImagePath) {
                    try {
                        await fs.unlink(tptappinglocImagePath);
                        console.log('🗑️ Cleaned up temporary TP_TAPPING_LOC image file');
                    } catch (cleanupError) {
                        console.warn('⚠️ Could not clean up temporary TP_TAPPING_LOC image:', cleanupError);
                    }
                }

                if (code === 0) {
                    console.log('✅ Python script completed successfully');

                    // Check if output file was created
                    try {
                        await fs.access(outputPath);
                        const fileStats = await fs.stat(outputPath);
                        console.log('📁 Output file created successfully');
                        console.log('📊 Output file size:', fileStats.size, 'bytes');
                        
                        // Return file download response
                        res.setHeader('Content-Disposition', `attachment; filename="${outputFilename}"`);
                        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.presentationml.presentation');
                        
                        const fileBuffer = await fs.readFile(outputPath);
                        console.log('📤 Sending file to client');
                        res.send(fileBuffer);
                        
                        // Clean up output file after sending
                        setTimeout(async () => {
                            try {
                                await fs.unlink(outputPath);
                                console.log('🗑️ Cleaned up output file');
                            } catch (cleanupError) {
                                console.warn('⚠️ Could not clean up output file:', cleanupError);
                            }
                        }, 5000); // 5 second delay to ensure download completes

                    } catch (error) {
                        console.error('❌ Output file not found or not accessible:', outputPath);
                        console.error('File access error:', error);
                        res.status(500).json({
                            success: false,
                            message: 'Proposal generation completed but output file not found'
                        });
                    }
                } else {
                    console.error('❌ Python script failed with code:', code);
                    res.status(500).json({
                        success: false,
                        message: `Proposal generation failed (code ${code}): ${stderr || 'Unknown error'}`
                    });
                }
            } catch (error) {
                console.error('❌ Error in Python process completion handler:', error);
                res.status(500).json({
                    success: false,
                    message: 'Internal server error during proposal generation'
                });
            }
        });

        pythonProcess.on('error', (error) => {
            console.error('❌ Python process error:', error);
            res.status(500).json({
                success: false,
                message: 'Failed to start Python script. Please ensure Python 3 is installed.'
            });
        });

    } catch (error) {
        console.error('❌ Unexpected error in proposal generation:', error);
        res.status(500).json({
            success: false,
            message: 'Unexpected server error'
        });
    }
});

// Health check endpoint
app.get('/api/health', (req, res) => {
    res.json({ 
        status: 'ok', 
        timestamp: new Date().toISOString(),
        message: 'TCNG Proposal Generator API is running'
    });
});

// Error handling middleware
app.use((error, req, res, next) => {
    if (error instanceof multer.MulterError) {
        if (error.code === 'LIMIT_FILE_SIZE') {
            return res.status(400).json({
                success: false,
                message: 'File too large. Maximum size is 10MB.'
            });
        }
    }
    
    console.error('Server error:', error);
    res.status(500).json({
        success: false,
        message: 'Internal server error'
    });
});

// Initialize directories and start server
createDirectories().then(() => {
    app.listen(PORT, () => {
        console.log('🚀 TCNG Proposal Generator Server Started');
        console.log(`📍 Server running at: http://localhost:${PORT}`);
        console.log(`📋 API endpoint: http://localhost:${PORT}/api/generate-proposal`);
        console.log(`💚 Health check: http://localhost:${PORT}/api/health`);
        console.log('🎯 Ready to generate proposals!');
        console.log('='*60);
    });
}).catch(error => {
    console.error('❌ Failed to initialize server:', error);
    process.exit(1);
});