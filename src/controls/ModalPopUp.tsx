//import { useState } from "react";
import * as React from 'react';
import styles from './ModalPopUp.module.scss';

interface ModalProps {
  message: string;
  onClose: any;
}
const Modal: React.FC<ModalProps> = ({ message, onClose }) => {
  return (
    <div className={styles.modalpopupContainer}>
      <div className={styles.modalPopUp}>
        <button className={styles.closeIcon} onClick={onClose}>
          &times;
        </button>
        <p>{message}</p>
        <button className={styles.closeButton} onClick={onClose}>
          Close
        </button>
      </div>
    </div>
  );
};

export default Modal;
