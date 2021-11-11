import React from 'react'
import { handleDropAsync } from '../utils/FileUtils'

const UploadFile = () => {
    return (
        <div>
            <input onDrop={handleDropAsync} type="file" id="fileUpload" />
        </div>
    )
}

export default UploadFile
