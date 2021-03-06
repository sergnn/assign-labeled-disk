import argparse
import sys
import win32api
import win32file


def get_drives_labels():
    drives = win32api.GetLogicalDriveStrings()
    used_drive_letters = {}
    for drive_letter in drives.split('\000')[:-1]:
        drive_letter = drive_letter.lower()
        try:
            used_drive_letters[drive_letter[0]] = win32api.GetVolumeInformation(drive_letter)[0].lower()
        except win32api.error:
            pass
    return used_drive_letters


def char_range(from_char, to_char):
    """Generates the characters from `from_char` to `to_char`, inclusive."""
    return {chr(c) for c in range(ord(from_char), ord(to_char) + 1)}


def assign_letter_by_label(label, letter):
    label_ = label.lower()
    letter_ = letter.lower()

    drives = get_drives_labels()

    if label_ in drives.values():
        search_drive = list(drives.keys())[list(drives.values()).index(label_)]
    else:
        print(f'Disk "{label}" not found.')
        return False

    if letter_ in drives and drives[letter_] == label_:
        print(f'Disk "{label}" already assigned to {letter.upper()}:\\')
        return True

    elif letter_ in drives:
        print(f'Letter {letter.upper()}:\\ already assigned to disk "{drives[letter_]}"')
        mounted_volume = win32file.GetVolumeNameForVolumeMountPoint(f'{letter_}:\\')
        win32file.DeleteVolumeMountPoint(f'{letter_}:\\')

        new_volume_letter = sorted(char_range('c', 'z') - set(drives.keys()))[0]

        print(f'Remapping partition {mounted_volume} to {new_volume_letter.upper()}:\\')
        win32file.SetVolumeMountPoint(f'{new_volume_letter}:\\', mounted_volume)

    elif search_drive:
        print(f'Disk "{label}" found as drive {search_drive.upper()}:\\')

    remap_from_ = f'{search_drive}:\\'
    remap_to_ = f'{letter_}:\\'
    print(f'Remapping from {remap_from_.upper()} to {remap_to_.upper()}')
    volume_name = win32file.GetVolumeNameForVolumeMountPoint(remap_from_)
    win32file.DeleteVolumeMountPoint(remap_from_)
    win32file.SetVolumeMountPoint(remap_to_, volume_name)

    return True


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Remaps labeled disk to selected letter')
    parser.add_argument('label', help='drive label')
    parser.add_argument('letter', help='new drive letter')
    args = parser.parse_args()
    if assign_letter_by_label(args.label, args.letter):
        print(f'Letter {args.letter.upper()}:\\ has been sucessfuly assigned to disk "{args.label}"')
    else:
        sys.exit(1)
